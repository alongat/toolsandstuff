require 'budgetbakers'
require 'date'
require 'awesome_print'
# CSV should be as follows in UTF-8
# Account Name, DATE,NAME,AMOUNT


module CsvToWallet
  MAP_FILENAME = 'CATMAP.csv'

  def self.push_csv_to_wallet(filename, email, apikey, params={})
    default_category = params[:default_category]
    api = Budgetbakers::API.new(email, apikey)
    load_name_category_map
    puts filename
    CSV.foreach(filename, { :col_sep => "," } ) do |row|
      account_name = row[0]
      date = row[1]
      name = row[2]
      amount = row[3]
      category = default_category || @category_map[name.downcase]
      if category.nil?
        puts "Enter category for #{name} | #{name.reverse}:"
        category = gets
        @category_map[name.downcase] = category.chomp!
      end
      body = { category_name: category,
               account_name: account_name,
               amount: amount,
               date: date,
               note: name }
      api.create_record(body)
    end
  rescue StandardError => e
    puts e
  ensure
    write_name_category_map unless @category_map&.empty?
  end

  def self.backup
    email = ENV['EMAIL']
    apikey = ENV['APIKEY']
    api = Budgetbakers::API.new(email, apikey)
    api.dump_records_to_file
  end

  def self.restore_from_backup(backup_file)
    email = ENV['EMAIL']
    apikey = ENV['APIKEY']
    api = Budgetbakers::API.new(email, apikey)
    CSV.foreach(backup_file, { :col_sep => ',' } ) do |row|
      puts row
      date = row[0]
      category = row[1] || 'Other'
      account_name = row[2]
      currency = row[3]
      type = row[4]
      amount = row[5]
      note = row[6]
      body = { category_name: category,
               account_name: account_name,
               amount: amount,
               date: date,
               note: note }

      api.create_record(body)
    end
  end

  def self.load_name_category_map(filename = MAP_FILENAME)
    @category_map = {}
    CSV.foreach(filename) { |row| @category_map[row[0].downcase] = row[1].downcase }
  rescue
    @category_map
  end

  def self.write_name_category_map(filename = MAP_FILENAME)
    CSV.open(filename, 'w') do |csv|
      @category_map.each { |x, y| csv << [x,y] }
    end
  end
end
