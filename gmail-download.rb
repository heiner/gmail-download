
require 'digest/sha1'
require 'fileutils'
require 'net/imap'
require 'zlib'

class GmailDownload
  include FileUtils

  attr_accessor :imap, :out, :slice

  ALL_MAIL = "[Google Mail]/All Mail"

  def initialize(compression=:gzip, slice=128, out=$stdout)
    @compression = compression
    @slice = slice
    @out = out
  end

  def overwrite_safe?(destdir)
    if File.exists?("labels")
      trashdir = File.join("labels", "Trash")
      require 'find'

      Find.find("labels") do |file|
        return false if not File.symlink?(file) and
                        not File.directory?(file) and
                        not File.dirname(file) == trashdir
      end
    end

    return true
  end

  def start_download(username, password, destdir, overwrite=false)
    cd destdir

    begin
      if overwrite and File.exists?("labels")
        rm_r "labels"
      end

      login(username, password)
      download
    rescue Interrupt
      out.print "Abort."
      imap.logout
    ensure
      out.puts ""
      imap.disconnect
    end
  end

  :private

  def log(str)
    $stderr.write(str) if $DEBUG
  end

  def progress(char)
    if (@column % 56) == 0
      @out.print("\n  ")
    elsif (@column % 4) == 0
      @out.print(" ")
    end
    @out.print(char)
    @out.flush
    @column += 1
  end

  def login(username, password)
    @imap = Net::IMAP.new('imap.gmail.com', 993, true)
    @imap.login(username + "@gmail.com", password)
  end

  def download
    labels = imap.list("", "*").map { |ml| ml.name }
    fail if not labels.include?(ALL_MAIL)

    mkdir "All Mail" unless File.exists?("All Mail")
    cd "All Mail"
    download_label(ALL_MAIL, true)
    cd ".."

    mkdir "labels" unless File.exists?("labels")
    cd "labels"
    labeldir = pwd

    labels -= [ALL_MAIL, "[Google Mail]"]
    labels.each do |label|
      out.print "\nHandling label #{label}"

      path = label.sub("[Google Mail]/", "")

      mkpath path
      cd path

      depth = path.count("/") + 2
      @all_mail_dir = File.join([".."]*depth, "All Mail")
      download_label(label)
      cd labeldir
    end
  end

  def download_label(label, allmail=false)
    @column = 0

    imap.examine(label)
    number_mails = imap.responses["EXISTS"][0]

    @out.print "All Mail contains #{number_mails} mails. " \
               "Downloading in blocks of #{@slice}." if allmail

    (1...number_mails).each_slice(@slice) do |numbers|
      ids_numbers = fetch_ids_numbers(numbers)

      ids_numbers.delete_if do |id, n|
        if already_there?(id.to_s)
          true
        elsif not allmail
          link_if_in_all_mail(id.to_s)
        end
      end

      if ids_numbers.empty?
        progress(".")
        next
      elsif ids_numbers.length < @slice
        progress(":")
      else
        progress("*")
      end

      log("Will get #{ids_numbers.inspect}\n#{numbers[0]}..#{numbers[-1]}\n")

      download_mails(ids_numbers)
    end
  end

  def fetch_ids_numbers(numbers)
    # log("#{numbers[0]}..#{numbers[-1]}\n")
    ids = imap.fetch(numbers, "BODY[HEADER.FIELDS (Message-ID)]")
    ids.map! { |fd| fd.attr.values.first }
    ids_numbers = ids.zip(numbers)

    idless_numbers = []
    ids_numbers.map! do |id, n|
      if id =~ /^Message-ID:\s+<?(\S+@[^\s>]+)>?\s*$/i
        [$1.gsub(/[\/]/, "_"), n] # make id filenameable
      else
        # Email without Message-ID?
        idless_numbers << n
        nil
      end
    end
    ids_numbers.compact!

    if not idless_numbers.empty?
      uids = imap.fetch(idless_numbers, "UID").map {|fd| fd.attr.values.first}
      ids_numbers.concat(uids.zip(idless_numbers))
    end

    return ids_numbers
  end

  def download_mails(ids_numbers)
    ids, numbers = ids_numbers.transpose

    mails = imap.fetch(numbers, "RFC822")
    mails.map! { |fd| fd.attr["RFC822"] }

    mails.zip(ids).each do |mail, id|
      # Handle email without Message-ID
      if id.is_a?(Integer)
        sha1 = Digest::SHA1.hexdigest(mail)
        link_if_in_all_mail(sha1)
        symlink(id2filename(sha1), id2filename(id.to_s))
        log("Email without Message-ID: #{id}\n")
        next if already_there?(sha1)
      end

      log("Write message #{id}\n")
      case @compression
      when :gzip
        Zlib::GzipWriter.open(id + ".gz") do |file|
          file.write(mail)
        end
      else
        File.open(id, 'w') do |file|
          file.write(mail)
        end
      end
    end
  end

  def already_there?(id)
    File.exists?(id + ".gz") or File.exists?(id)
  end

  def in_all_mail?(id)
    fn = File.join(@all_mail_dir, id)
    [fn + ".gz", fn].each do |p|
      return p if File.exists?(p)
    end
    return false
  end

  def link_if_in_all_mail(id)
    path = in_all_mail?(id)
    if path
      extname = File.extname(path)
      extname = "" if not extname == ".gz"
      id += extname
      if File.exists?(id)
        return true if File.symlink?(id) and File.readlink(id) == path
        log("File already exists: #{id}, but doesn't point to #{path}")
      end
      symlink(path, id)
      return true
    end
  end

  def id2filename(id)
    case @compression
    when :gzip
      id + ".gz"
    else
      id
    end
  end

end

if __FILE__ == $0
  require "getoptlong"
  require 'yaml'

  require 'rubygems'
  require 'highline/import'

  opts = GetoptLong.new(
                        [ '-u', GetoptLong::NO_ARGUMENT ],
                        [ '-D', GetoptLong::NO_ARGUMENT ]
                        )
  config = { :overwrite=>false, :compression=>:gzip }
  opts.each do |opt, arg|
    case opt
    when '-D'
      config[:overwrite] = true
    when '-u'
      config[:compression] = nil
    end
  end

  gmail = GmailDownload.new(config[:compression])

  config_filename = File.expand_path("~/.gmail-downloadrc.yaml")
  if File.exists?(config_filename)
    config.merge!(YAML.load_file(config_filename))
  end

  begin
    username = ask("Gmail user: " ) do |q|
      if config.has_key?("username")
        q.default = config["username"]
      end
    end
    password = ask("Password: ") { |q| q.echo = false }
    workdir = ask("Directory: ") do |q|
      if config.has_key?("directory")
        q.default = config["directory"]
      else
        current = pwd
        if File.basename(current) == "gmail-emails"
          q.default = current
        else
          q.default = File.join(current, "gmail-emails")
        end
      end
    end
  rescue Interrupt
    exit 1
  end

  if config[:overwrite] and not gmail.overwrite_safe?(workdir)
    puts "\nWARNING: There are non-symlink files outside of labels/Trash."
    puts "(Use 'find #{File.join(workdir, 'labels')} -type f' to see them)"
    yn = agree("Still overwrite labels/ directory? ")
    unless yn
      puts "Abort."
      exit
    end
  end

  File.open(config_filename, "w") do |file|
    YAML.dump({"username"=>username, "directory"=>workdir}, file)
  end

  gmail.start_download(username, password, workdir, config[:overwrite])
end
