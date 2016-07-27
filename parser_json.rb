class  ImportRaytingWorker
  require 'rubyXL'
  require 'json'

  def self.perform


    debug_file = open("output.txt", 'w')
    result_file = open("result.json", 'w')
    debug_file.write("Begin\n")
    workbook = RubyXL::Parser.parse("rating_2016.xlsm")
    worksheet = workbook.worksheets[0]
    total = 0
    worksheet.each { |row| total +=1}
    debug_file.write("total is #{total}\n" )
    index = 1
    cell = worksheet.sheet_data[index][0]
    result = []



    debug_file.write("Before while\n")
    while index < total
      #sheet_data = worksheet.sheet_data[index]
      cell = worksheet.sheet_data[index][0] if worksheet.sheet_data[index]
      current  = cell && cell.value

      if current == "Институт / факультет"

        cell = worksheet.sheet_data[index][5] if worksheet.sheet_data[index]
        current_inst  = cell && cell.value
        debug_file.write("Institute: #{current_inst}\n")

	result_inst  = Hash.new
	result_inst[:institute_title] = current_inst
	result_inst[:specialities] = []
	result.push(result_inst)


        index +=1
        next
      elsif current == "Направление подготовки / специальность"
        cell = worksheet.sheet_data[index][5] if worksheet.sheet_data[index]
        spec_title  = cell && cell.value
        debug_file.write("speciality: #{spec_title}\n")
        spec = Hash.new
	spec[:title] = spec_title

        index +=2
        cell = worksheet.sheet_data[index][1] if worksheet.sheet_data[index]
        description  = cell && cell.value
        debug_file.write("set: #{description}\n")
	spec[:description] = description

        index +=1
        cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
        plan  = cell && cell.value
        debug_file.write("plan: #{plan}\n")
	spec[:plan] = plan

        index +=1
        cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
        submitted  = cell && cell.value
        debug_file.write("submitted: #{submitted}\n")
	spec[:submitted] = submitted

        index +=1
        cell = worksheet.sheet_data[index][2] if worksheet.sheet_data[index]
        contest = cell && cell.value
        debug_file.write("contest: #{contest}\n")
	spec[:contest] = contest

        index +=4
        cell = worksheet.sheet_data[index][7] if worksheet.sheet_data[index]
        exam_1 = cell && cell.value
  spec[:exam_1] = exam_1
        cell = worksheet.sheet_data[index][8] if worksheet.sheet_data[index]
        exam_2 = cell && cell.value
  spec[:exam_2] = exam_2
        cell = worksheet.sheet_data[index][9] if worksheet.sheet_data[index]
        exam_3 = cell && cell.value
  spec[:exam_3] = exam_3
        cell = worksheet.sheet_data[index][10] if worksheet.sheet_data[index]
        exam_4 = cell && cell.value
  spec[:exam_4] = exam_4


        index +=1
        cell = worksheet.sheet_data[index][1] if worksheet.sheet_data[index]
        rating = cell && cell.value

        if rating == "Рейтинг"
          index +=1
	  spec[:ratings] = []
    spec[:names]   = []
	  index = get_set_ratings(index, worksheet,  debug_file, spec)
	  result_inst[:specialities].push(spec)
	  result.pop
	  result.push(result_inst)

        end
      end

      index += 1
    end
    debug_file.close
    result_file.write(result.to_json  )
    result_file.close
  end

  def self.get_set_ratings(index, worksheet, debug_file, spec)
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
	  rating = Hash.new
	  rating[:title] = title
	  rating[:list]  = []

	  index +=1
          index = get_rating_data(index, worksheet,  debug_file, rating, spec)
	  spec[:ratings].push(rating)


        else
          index +=1
        end
      end
      i +=1
    end
    index
  end
  def self.get_rating_data(index, worksheet, debug_file, rating, spec)
        until_cell = worksheet.sheet_data[index][3]
    while until_cell
        submit = Hash.new
        cell = worksheet.sheet_data[index][2]
        position  = cell && cell.value
        submit[:position] = position

        cell = worksheet.sheet_data[index][3]
        id  = cell && cell.value
        submit[:id] = id

        cell = worksheet.sheet_data[index][4]
        name  = cell && cell.value
        submit[:name] = name
        last_name = name.split.first
        spec[:names].push(last_name)

        cell = worksheet.sheet_data[index][5]
        total  = cell && cell.value
        submit[:total] = total

        cell = worksheet.sheet_data[index][5]
        achiev = cell && cell.value
        submit[:achiev] = achiev

        cell = worksheet.sheet_data[index][7] if worksheet.sheet_data[index]
        exam_1 = cell && cell.value
        submit[:exam_1] = exam_1

        cell = worksheet.sheet_data[index][8] if worksheet.sheet_data[index]
        exam_2 = cell && cell.value
        submit[:exam_2] = exam_2

        cell = worksheet.sheet_data[index][9] if worksheet.sheet_data[index]
        exam_3 = cell && cell.value
        submit[:exam_3] = exam_3

        cell = worksheet.sheet_data[index][10] if worksheet.sheet_data[index]
        exam_4 = cell && cell.value
        submit[:exam_4] = exam_4

        cell = worksheet.sheet_data[index][11] if worksheet.sheet_data[index]
        doc_original = cell && cell.value
        submit[:doc_original] = doc_original


        cell = worksheet.sheet_data[index][11] if worksheet.sheet_data[index]
        privilege = cell && cell.value

        submit[:privilege] = privilege

        cell = worksheet.sheet_data[index][12] if worksheet.sheet_data[index]
        benefit = cell && cell.value
        submit[:benefit] = benefit

        debug_file.write("student: #{position}, #{id}, #{name}, #{total} \n")
	      rating[:list].push(submit)
        until_cell = nil
        index +=1
        until_cell = worksheet.sheet_data[index][3] if worksheet.sheet_data[index]
    end
    index
  end

end
ImportRaytingWorker.perform
