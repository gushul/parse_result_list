require 'rubyXL'


output_file = open("output.txt", 'w')
workbook = RubyXL::Parser.parse("rating_2015.xlsm")
worksheet = workbook.worksheets[0]
index = 1

cell = worksheet.sheet_data[index][0]
if !cell || !cell.value
  output_file.write("Ошибка: не указан PKUID\n")
else
  output_file.write(cell.value)
end

output_file.close
