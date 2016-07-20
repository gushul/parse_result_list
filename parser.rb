class  ImportRaytingWorker
	require 'rubyXL'
	require 'json'


	def self.rating_total
		debug_file = open("output.txt", 'w')
		result_file = open("result.txt", 'w')
		debug_file.write("Begin\n")
		workbook = RubyXL::Parser.parse("rating_2015.xlsm")
		worksheet = workbook.worksheets[0]
		total = 0
		worksheet.each { |row| total +=1}
		debug_file.write("total is #{total}\n" )
		index = 1
		print total
		print "\n"

		rating_total = 0
		debug_file.write("Before while\n")
		while index < total
		    cell = worksheet.sheet_data[index][0] if worksheet.sheet_data[index]
		    current = cell && cell.value
		    #rating = cell && cell.value
			    if current == "Направление подготовки / специальность"

			 # if rating == "Рейтинг"
				  index +=1
				  rating_total +=1
				print "================\n"

				print "\n"
				print "#{index}: #{current}: #{rating_total}"
				print "\n"
				print "================\n"
			  end
			  index +=1
		end
		print total
	end

	def self.perform


		debug_file = open("output.txt", 'w')
		result_file = open("result.txt", 'w')
		debug_file.write("Begin\n")
		workbook = RubyXL::Parser.parse("rating_2015.xlsm")
		worksheet = workbook.worksheets[0]
		total = 0
		worksheet.each { |row| total +=1}
		debug_file.write("total is #{total}\n" )
		index = 1
		cell = worksheet.sheet_data[index][0]



		debug_file.write("Before while\n")
		while index < total
		  #sheet_data = worksheet.sheet_data[index]
		  cell = worksheet.sheet_data[index][0] if worksheet.sheet_data[index]
		  results = Hash.new
		  current  = cell && cell.value

		  if current == "Институт / факультет"
		    cell = worksheet.sheet_data[index][5] if worksheet.sheet_data[index]
		    current_inst  = cell && cell.value
		    debug_file.write("Institute: #{current_inst}\n")
		    results.store(current_inst, "")
		    spec_hash = Hash.new


		    index +=1
		    next
		  elsif current == "Направление подготовки / специальность"
		    cell = worksheet.sheet_data[index][5] if worksheet.sheet_data[index]
		    current_spec  = cell && cell.value
		    debug_file.write("speciality: #{current_spec}\n")
		    spec_hash.store(current_spec, "")
		    spec_set = Hash.new

		    index +=2
		    cell = worksheet.sheet_data[index][1] if worksheet.sheet_data[index]
		    spec_set_name  = cell && cell.value
		    debug_file.write("set: #{spec_set_name}\n")

		    index +=1
		    cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
		    plan  = cell && cell.value
		    debug_file.write("plan: #{plan}\n")

		    index +=1
		    cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
		    submitted  = cell && cell.value
		    debug_file.write("submitted: #{submitted}\n")

		    index +=1
		    cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
		    contest = cell && cell.value
		    debug_file.write("contest: #{contest}\n")

		    index +=5
		    cell = worksheet.sheet_data[index][1] if worksheet.sheet_data[index]
		    rating = cell && cell.value

			  if rating == "Рейтинг"
				  index +=1
				index = get_set_ratings(index, worksheet,  debug_file)
			  end
		  end

		  index += 1
		end
		debug_file.close
	end

	def self.get_set_ratings(index, worksheet, debug_file)
		i = 0
		rate_titles = ["Без вступительных испытаний", "Общий конкурс"]
		debug_file.write("Рейтинг:\n")
		while i < 2
			cell = worksheet.sheet_data[index][2]
			title  = cell && cell.value
			if worksheet.sheet_data[index+1]
				cell = worksheet.sheet_data[index+1][2]
				next_cell  = cell && cell.value
				if(rate_titles.include?(title) && next_cell.is_a?(Integer))
				  debug_file.write("#{title}:\n")
				  index +=1
				  index = get_rating_data(index, worksheet,  debug_file)
				else
				  index +=1
				end
			end
			i +=1
		end
		index
	end
	def self.get_rating_data(index, worksheet, debug_file)
	      until_cell = worksheet.sheet_data[index][3]
	  while until_cell
	      cell = worksheet.sheet_data[index][2]
	      rate  = cell && cell.value
	      cell = worksheet.sheet_data[index][3]
	      id  = cell && cell.value
	      cell = worksheet.sheet_data[index][4]
	      name  = cell && cell.value
	      cell = worksheet.sheet_data[index][5]
	      sum  = cell && cell.value
	      #result_file.write("student: #{rate}, #{id}, #{name}, #{sum} \n")
	      debug_file.write("student: #{rate}, #{id}, #{name}, #{sum} \n")
	      until_cell = nil
	      index +=1
	      until_cell = worksheet.sheet_data[index][3] if worksheet.sheet_data[index]
	  end
	  index
	end

end
ImportRaytingWorker.perform
#ImportRaytingWorker.rating_total

