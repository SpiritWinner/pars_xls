require 'spreadsheet'
=begin
count = 2
loop do
=end
workbook = Spreadsheet.open "C:/Users/Spirit/Desktop/5/1_2.xls"
worksheets = workbook.worksheets
# puts "Found #{worksheets.count} worksheets"
worksheets.each do |worksheet|
  # puts "Reading: #{worksheet.name}"
  worksheet.rows.each do |row|
    row_cells = row.to_a.map { |v| v.methods.include?(:value) ? v.value : v }
    puts (row_cells[6]).to_s if !row_cells[6].nil? && row_cells[6] != "E-mail"
  end
  # puts "Read #{num_rows} rows"
end
=begin
        count += 1
        break if count == 4
=end
# end
