require 'rubyXL'
require 'json'


output_file = open("output.txt", 'w')
result_file = open("result.txt", 'w')
output_file.write("Begin\n")
workbook = RubyXL::Parser.parse("rating_2015.xlsm")
worksheet = workbook.worksheets[0]
total = 0
worksheet.each { |row| total +=1}
output_file.write("total is #{total}\n" )
index = 1
cell = worksheet.sheet_data[index][0]



output_file.write("Before while\n")
#while !cell
while index < total
  #sheet_data = worksheet.sheet_data[index]
  cell = worksheet.sheet_data[index][0] if worksheet.sheet_data[index]
  results = Hash.new
  current  = cell && cell.value

  if current == "Институт / факультет"
    cell = worksheet.sheet_data[index][5] if worksheet.sheet_data[index]
    current_inst  = cell && cell.value
    output_file.write("Institute: #{current_inst}\n")
    results.store(current_inst, "")
    spec_hash = Hash.new

    index +=1
    cell = worksheet.sheet_data[index][5] if worksheet.sheet_data[index]
    current_spec  = cell && cell.value
    output_file.write("speciality: #{current_spec}\n")
    spec_hash.store(current_spec, "")
    spec_set = Hash.new

    index +=2
    cell = worksheet.sheet_data[index][1] if worksheet.sheet_data[index]
    spec_set_name  = cell && cell.value
    output_file.write("set: #{spec_set_name}\n")

    index +=1
    cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
    plan  = cell && cell.value
    output_file.write("plan: #{plan}\n")

    index +=1
    cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
    submitted  = cell && cell.value
    output_file.write("submitted: #{submitted}\n")

    index +=1
    cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
    contest = cell && cell.value
    output_file.write("contest: #{contest}\n")

    index +=5
    cell = worksheet.sheet_data[index][1] if worksheet.sheet_data[index]
    rating = cell && cell.value

    output_file.write("#{rating}:\n")
	  if rating == "Рейтинг"
	    index +=2
	    cell = worksheet.sheet_data[index][2]
	    title  = cell && cell.value

	    output_file.write("#{title}:\n")

	    while
	      index += 1
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
