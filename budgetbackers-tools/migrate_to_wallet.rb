require './collect_excels_to_csv'
require './csv-to-wallet'

MigrateXLToCsv.do_the_trick('in','out')

Dir.foreach('out') do |item|
  next unless item.include?('.csv')
  puts 'Converting ' + item
  CsvToWallet.push_csv_to_wallet('out/' + item, email, apikey)
end