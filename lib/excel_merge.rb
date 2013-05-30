require "spreadsheet"
require "csv"


class ExcelMerge
  TEMPLATE_OFFSET = 9

  class << self

    def merge_template_with_csv(template, csv)
      template = Spreadsheet.open(template)
      csv      = CSV.parse(csv)

      worksheet = template.worksheets.first

      headers = csv.shift

      csv.each_with_index do |csv_row, row_index|

        csv_row.each_with_index do |csv_cell, col_index|
          next if csv_cell.nil?
          worksheet[row_index + TEMPLATE_OFFSET, col_index] = csv_cell
        end
      end

      template
    end
  end
end
