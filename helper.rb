class Helper

  @@column_widths = {}

  def self.convert_letters_to_index(index_name)
    iteration_index = 0
    column_index = 0
    index_name.each_char do |char|
      column_index = column_index + (char.ord - 'A'.ord)*(26**iteration_index)
      iteration_index = iteration_index + 1
    end
    column_index
  end

  def self.empty_cell?(sheet, row, col)
    sheet.nil? || sheet[row].nil? || sheet[row][col].nil? || sheet[row][col].value.nil?
  end

  def self.cell_value(sheet, row, col)
    return nil if sheet.nil? || sheet[row].nil? || sheet[row][col].nil?
    sheet[row][col].value
  end

  def self.coordinates_to_sheet_and_cell(coordinates_with_sheet)
    coordinates_with_sheet.split ':'
  end

  def self.string_coordinates_to_indexes(string_coordinates)
    column = ''
    string_coordinates.each_char do |char|
      char =~ /[A-Z]/ ? column << char : break
    end

    row = string_coordinates[column.size .. -1]

    [Helper.convert_letters_to_index(column), Integer(row).to_i - 1, column, row]
  end

  def self.split_combo_rule_in_head_and_tail(combo_string)
    first_rule_end = combo_string.index('∩') || combo_string.size
    first_rule = combo_string[0..first_rule_end-1]
    second_rule_combo = combo_string[first_rule_end+1..-1]
    [first_rule, second_rule_combo]
  end

  def self.set_cell_data(sheet, row_idx, col_idx, value, opts={change_col_width: true})
    if opts[:change_col_width]
      width = @@column_widths[col_idx]
      value_size = value.to_s.size + 1
      if width.nil?
        sheet.change_column_width col_idx, (@@column_widths[col_idx] = value_size)
      elsif width < value_size
        sheet.change_column_width col_idx, (@@column_widths[col_idx] = value_size)
      end
    end

    if (sheet[row_idx].nil? || (cell = sheet[row_idx][col_idx]).nil?)
      sheet.add_cell row_idx, col_idx, value.to_s
    else
      cell.change_contents value.to_s
    end

    if fill_color = opts[:fill_color]
      #worksheet.sheet_data[0][0].change_fill('0ba53d')
      sheet.sheet_data[row_idx][col_idx].change_fill fill_color
    end
    cell
  end

  def self.array_intersection(array1, array2)
    return [] if (array1.empty? || array2.empty?) || (array1.first > array2.last) || (array1.last < array2.first)

    intersection = []
    array1.each do |a1|
      return intersection if array2.empty? || a1 > array2.last
      array2.each do |a2|
        if a1 == a2
          intersection << a2
          break
        end
      end
    end
    intersection
  end
end