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
    diff_cell = sheet.cell('A', 1)
    puts diff_cell
    if diff_cell == 'פירוט עסקאות בכרטיסים'
      # Discount Bank Credit Cards
      account_name = 'Credit Cards'
      date_col = 1 ; name_col = 2 ; amount_col = 6 ; start_row = 17
      process_csv(account_name, amount_col, csv, date_col, name_col, num_of_rows, sheet, start_row)
    elsif diff_cell == 'תנועות אחרונות'
      # Discount Bank
      account_name = 'Bank account'
      date_col = 1 ; name_col = 2 ; amount_col = 3 ; start_row = 11

      process_csv(account_name, amount_col, csv, date_col, name_col, num_of_rows, sheet, start_row, false)
    elsif diff_cell.include? 'פירוט עסקות נכון לתאריך:'
      # Visa Cal Credit Cards
      account_name = 'Credit Cards'
      date = '02/' + (sheet.cell('A', 2).scan(/[0-1][0-9]\/[0-9]{2}(?:[0-9]{2})?/).first)
      name_col = 1 ; amount_col = 3 ; start_row = 4

      (start_row..num_of_rows-1).each do |row|
        amount, _, sheet_row = extract_cells(amount_col, 3, row, sheet)
        next if amount.zero?
        create_csv_line(account_name, -amount, csv, date, name_col, sheet_row)
      end
    elsif diff_cell.include? 'תאריך עסקה'
      # Leumi Card Credit Cards
      account_name = 'Credit Cards'
      name_col = 2 ; amount_col = 6 ; date_col = 1 ; start_row = 2

      process_csv(account_name, amount_col, csv, date_col, name_col, num_of_rows, sheet, start_row)
    elsif diff_cell.include? 'המשתמשים'
      # Max (leumi card) Credit Cards
      account_name = 'Credit Cards'
      name_col = 2 ; amount_col = 5 ; start_row = 5
      date = '02/' + (sheet.cell('A', 3).scan(/[0-1][0-9]\/[0-9]{2}(?:[0-9]{2})?/).first)

      (0...xlsx.sheets.count).each do |i|
        sheet = xlsx.sheet(i)
        num_of_rows = sheet.last_row
        (start_row..num_of_rows).each do |row|
          amount, _, sheet_row = extract_cells(amount_col, 1, row, sheet)
          next if amount.zero?
          create_csv_line(account_name, -amount, csv, date, name_col, sheet_row)
        end
      end
    end
  end

  private

  def self.process_csv(account_name, amount_col, csv, date_col, name_col, num_of_rows, sheet, start_row, flip_amount=true)
    (start_row..num_of_rows).each do |row|
      amount, date, sheet_row = extract_cells(amount_col, date_col, row, sheet)
      amount = -amount if flip_amount
      next if amount.zero?
      create_csv_line(account_name, amount, csv, date, name_col, sheet_row)
    end
  end

  def self.extract_cells(amount_col, date_col, row, sheet)
    sheet_row = sheet.row(row)
    amount = sheet_row[amount_col].to_s.gsub(/,|₪|\s/, '').to_f
    date = sheet_row[date_col].to_s
    return amount, date, sheet_row
  end

  def self.create_csv_line(account_name, amount, csv, date, name_col, sheet_row)
    csv << [account_name,
            date,
            clean_name(sheet_row[name_col].to_s),
            amount]
  end
end
