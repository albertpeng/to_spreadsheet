require 'rubyXL'
require 'rubyXL/writer/worksheet_writer'

module ToSpreadsheet
  module XLS
    extend self

    def to_io(html)
      workbook = RubyXL::Workbook.new
      workbook.worksheets= []
      Nokogiri::HTML::Document.parse(html).css('table').each_with_index do |xml_table, i|
        workbook.worksheets << (worksheet = RubyXL::Worksheet.new(workbook, xml_table.css('caption').inner_text.presence || "Sheet #{i + 1}"))
        xml_table.css('tr').each_with_index do |row_node, row|
          row_node.css('th,td').each_with_index do |col_node, col|
            node = typed_node_val(col_node)
            cell = worksheet.add_cell(row, col, node)
          end
        end
      end
      tmp_file = File.new(%w[to_spreadsheet .xlsx].join, 'wb')
      workbook.write(tmp_file.path)
      StringIO.new(File.open(tmp_file.path).binmode.read)
    end

    private

    def typed_node_val(node)
      val = node.inner_text
      case node[:class]
        when /decimal|float/
          val.to_f
        when /num|int/
          val.to_i
        when /datetime/
          DateTime.parse(val)
        when /date/
          Date.parse(val)
        else
          val
      end
    end
  end
end
