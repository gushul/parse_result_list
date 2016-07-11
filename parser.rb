require 'rubyXL'


output_file = open("output.txt", 'w')
output_file.write("Begin\n")
workbook = RubyXL::Parser.parse("mini_rating_2015.xlsm")
worksheet = workbook.worksheets[0]
index = 1
cell = worksheet.sheet_data[index][0]



output_file.write("Before while\n")
#while !cell
while index < 100
  output_file.write("in while #{index}\n")
  sheet_data = worksheet.sheet_data[index]
  rating  = sheet_data[1] && sheet_data[1].value if sheet_data
  if rating == "Рейтинг"
    output_file.write("#{rating}\n")
    index +=2
    
    sheet_data = worksheet.sheet_data[index]
    id = sheet_data[3].value
    while id
      output_file.write("in while with id \n")
      output_file.write("id_value:#{id.value} \n")
      index += 1
      sheet_data = worksheet.sheet_data[index]
      id = sheet_data[3]
    end
  end
  index += 1
end

output_file.close
