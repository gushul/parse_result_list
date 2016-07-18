require 'rubyXL'
require 'json'


output_file = open("output.txt", 'w')
result_file = open("result.txt", 'w')
output_file.write("Begin\n")
workbook = RubyXL::Parser.parse("rating_2015.xlsm")
worksheet = workbook.worksheets[0]
index = 1
cell = worksheet.sheet_data[index][0]



output_file.write("Before while\n")
#while !cell
while index < 200
  output_file.write("in while #{index}\n")
  #sheet_data = worksheet.sheet_data[index]
  cell = worksheet.sheet_data[index][0] if worksheet.sheet_data[index]
  results = Hash.new
  current  = cell && cell.value

  if current == "Институт / факультет"
    cell = worksheet.sheet_data[index][5] if worksheet.sheet_data[index]
    current_inst  = cell && cell.value
    output_file.write("#{current_inst}\n")
    results.store(current_inst, "")
    spec_hash = Hash.new

    index +=1
    cell = worksheet.sheet_data[index][5] if worksheet.sheet_data[index]
    current_spec  = cell && cell.value
    output_file.write("#{current_spec}\n")
    spec_hash.store(current_spec, "")
    spec_set = Hash.new

    index +=2
    cell = worksheet.sheet_data[index][1] if worksheet.sheet_data[index]
    spec_set_name  = cell && cell.value 
    output_file.write("#{spec_set_name}\n")




    if current == "Рейтинг"
      output_file.write("#{current}:\n")
      index +=2
      cell = worksheet.sheet_data[index][3]
      id  = cell && cell.value
      output_file.write("rayt +2#{id}\n")

      while index < 200
        index += 1
        output_file.write("index is #{index}\n")
        if worksheet.sheet_data[index]
          cell = worksheet.sheet_data[index][2]
          rate  = cell && cell.value
          cell = worksheet.sheet_data[index][3]
          id  = cell && cell.value
          cell = worksheet.sheet_data[index][4]
          name  = cell && cell.value
          cell = worksheet.sheet_data[index][5]
          sum  = cell && cell.value
          result_file.write("student: #{rate}, #{id}, #{name}, #{sum} \n")
          output_file.write("student: #{rate}, #{id}, #{name}, #{sum} \n")
        else
     break
        end
      end
    end
  end

  index += 1
end

output_file.close
