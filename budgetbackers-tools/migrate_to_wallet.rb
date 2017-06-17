require '../budgetbackers-tools/collect_excels_to_csv'
require '../budgetbackers-tools/csv-to-wallet'

MigrateXLToCsv.do_the_trick('in','out')

email = ENV['EMAIL']
apikey = ENV['APIKEY']

Dir.foreach('out') do |item|
  next unless item.include?('.csv')
  puts 'Converting ' + item
  CsvToWallet.push_csv_to_wallet('out/' + item, email, apikey)
end