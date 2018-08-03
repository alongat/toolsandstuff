require 'roo'

module MigrateXLToCsv
  def self.do_the_trick(in_dir, out_dir)
    Dir.foreach(in_dir) do |item|
      next unless item.include?('.xlsx')
      puts 'Converting ' + item
      CSV.open(out_dir + '/' + item.sub('.xlsx','.csv'), 'w', encoding: 'utf-8') do |csv|
        create_csvs_from_xls(in_dir + '/' + item, csv)
      end
    end
  end

  def self.clean_name(name)
    name.gsub(/'|"|,/,'')
  end


  def self.create_csvs_from_xls(in_xls, csv)
    xlsx = Roo::Spreadsheet.open(in_xls)
    num_of_rows = xlsx.sheet(0).last_row
    sheet = xlsx.sheet(0)
    if sheet.cell('A',1) == 'פירוט עסקאות בכרטיסים'
      # Discount Bank Credit Cards
      date_col = 1
      name_col = 2
      amount_col = 6
      start_row = 17
      account_name = 'Credit Cards'

      (start_row..num_of_rows).each do |row|
        sheet_row = sheet.row(row)
        csv << [account_name,
                sheet_row[date_col].to_s,
                clean_name(sheet_row[name_col].to_s),
                -sheet_row[amount_col].to_f]
      end
    elsif sheet.cell('A',1) == 'תנועות אחרונות'
      # Discount Bank
      account_name = 'Bank account'
      date_col = 1
      name_col = 2
      amount_col = 3
      start_row = 11

      (start_row..num_of_rows).each do |row|
        sheet_row = sheet.row(row)
        csv << [account_name,
                sheet_row[date_col].to_s,
                clean_name(sheet_row[name_col].to_s),
                sheet_row[amount_col].to_f]
      end
    elsif sheet.cell('A',1).include? 'פירוט עסקות נכון לתאריך:'
      # Visa Cal Credit Cards
      date = '02/' + (sheet.cell('A', 2).scan(/[0-1][0-9]\/[0-9]{2}(?:[0-9]{2})?/).first)
      name_col = 1
      amount_col = 3
      start_row = 4
      account_name = 'Credit Cards'

      (start_row..num_of_rows).each do |row|
        sheet_row = sheet.row(row)
        amount = sheet_row[amount_col].to_s.gsub(/,|₪|\s/, '')
        amount = -amount.to_f

        csv << [account_name,
                date,
                clean_name(sheet_row[name_col].to_s),
                amount]
      end
    elsif sheet.cell('A',1).include? 'תאריך עסקה'
      # Leumi Card Credit Cards
      name_col = 2
      amount_col = 6
      date_col = 1
      start_row = 2
      account_name = 'Credit Cards'

      (start_row..num_of_rows).each do |row|
        sheet_row = sheet.row(row)
        amount = sheet_row[amount_col].to_s.gsub(/,|₪|\s/, '').to_f
        next if amount.zero?
        csv << [account_name,
                sheet_row[date_col],
                clean_name(sheet_row[name_col].to_s),
                -amount]
      end
    end
  end
end
