require 'rubyXL'

#monkeypatching rubyXL, pull request submitted: https://github.com/gilt/rubyXL/pull/47
module RubyXL
  class Workbook
    def is_date_format?(num_fmt)
      num_fmt.downcase!
      skip_chars = ['$', '-', '+', '/', '(', ')', ':', ' ']
      num_chars = ['0', '#', '?']
      non_date_formats = ['0.00e+00', '##0.0e+0', 'general', '@']
      date_chars = ['y','m','d','h','s']

      state = 0
      s = ''
      num_fmt.split(//).each do |c|
        if state == 0
          if c == '"'
            state = 1
          elsif ['\\', '_', '*'].include?(c)
            state = 2
          elsif skip_chars.include?(c)
            next
          else
            s << c
          end
        elsif state == 1
          if c == '"'
            state = 0
          end
        elsif state == 2
          state = 0
        end
      end
      s.gsub!(/\[[^\]]*\]/, '')
      if non_date_formats.include?(s)
        return false
      end
      separator = ';'
      got_sep = 0
      date_count = 0
      num_count = 0
      s.split(//).each do |c|
        if date_chars.include?(c)
          date_count += 1
        elsif num_chars.include?(c)
          num_count += 1
        elsif c == separator
          got_sep = 1
        end
      end
      if date_count > 0 && num_count == 0
        return true
      elsif num_count > 0 && date_count == 0
        return false
      elsif date_count
        # ambiguous result
      elsif got_sep == 0
        # constant result
      end
      return date_count > num_count
    end
  end
end
# end monkey patch submitted

# other monkey patch
module RubyXL
  class Cell
    def number_format
      if !@value.is_a?(String)
        if @workbook.num_fmts_by_id
          num_fmt_id = xf_id()[:numFmtId]
          tmp_num_fmt = @workbook.num_fmts_by_id[num_fmt_id]
          return (tmp_num_fmt &&tmp_num_fmt[:attributes] && tmp_num_fmt[:attributes][:formatCode]) ? tmp_num_fmt[:attributes][:formatCode] : nil
        end
      end
    end
    def fill_color
      if !@value.is_a?(String)
        if @workbook.num_fmts_by_id
          num_fmt_id = xf_id()[:numFmtId]
          tmp_num_fmt = @workbook.num_fmts_by_id[num_fmt_id]
          return (tmp_num_fmt &&tmp_num_fmt[:attributes] && tmp_num_fmt[:attributes][:formatCode]) ? tmp_num_fmt[:attributes][:formatCode] : nil
        end
      end
    end
    
  end
end
# end of monkey patch 

module Workbook
  module Readers
    module XlsxReader
      def load_xlsx file_obj
        file_obj = file_obj.path if file_obj.is_a? File
        sp = RubyXL::Parser.parse(file_obj)
        template.add_raw sp
        parse_xlsx sp
      end
      
      def parse_xlsx xlsx_spreadsheet=template.raws[RubyXL::Workbook], options={}
        options = {:additional_type_parsing=>false}.merge options
        #number_of_worksheets = xlsx_spreadsheet.worksheets.count
        xlsx_spreadsheet.worksheets.each_with_index do |worksheet, si|
          s = create_or_open_sheet_at(si)    
          col_widths = xlsx_spreadsheet.worksheets.first.cols.collect{|a| a[:attributes][:width].to_f if a[:attributes]}
          worksheet.each_with_index do |row, ri|
            r = s.table.create_or_open_row_at(ri)
            
            row.each_with_index do |cell,ci|
              if cell
                r[ci] = Workbook::Cell.new nil
              else
                r[ci] = Workbook::Cell.new cell.value 
                
                r[ci].parse!    
                
                xls_format = cell.style_index
                col_width = nil
              
                if ri == 0
                  col_width = col_widths[ci]
                end
                f = template.create_or_find_format_by "style_index_#{cell.style_index}", col_width
                f[:width]= col_width
                f[:background_color] = "##{cell.fill_color}"
                f[:number_format] = ms_formatting_to_strftime(cell.number_format)
                f[:font_family] = cell.font_name
                f[:color] = "##{cell.font_color}"

                f.add_raw xls_format
     
                r[ci].format = f
              end
            end
          end
        end
      end
    private 
      def ms_formatting_to_strftime ms_nr_format   
        if ms_nr_format    
          ms_nr_format = ms_nr_format.downcase
          return nil if ms_nr_format == 'general'
          ms_nr_format.gsub('yyyy','%Y').gsub('dddd','%A').gsub('mmmm','%B').gsub('ddd','%a').gsub('mmm','%b').gsub('yy','%y').gsub('dd','%d').gsub('mm','%m').gsub('y','%y').gsub('%%y','%y').gsub('d','%e').gsub('%%e','%d').gsub('m','%m').gsub('%%m','%m').gsub(';@','').gsub('\\','')
        end
      end
    end
  end
end
