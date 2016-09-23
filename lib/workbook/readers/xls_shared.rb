# -*- encoding : utf-8 -*-

require 'date'

module Workbook
  module Readers
    module XlsShared

      # Converts standard (ruby/C++/unix/...) strftime formatting to MS's formatting
      #
      # @param [String, nil] ms_nr_format (nil returns nil)
      # @return [String, nil]
      def ms_formatting_to_strftime ms_nr_format
        ms_nr_format = num_fmt_id_to_ms_formatting(ms_nr_format) if ms_nr_format.is_a? Integer
        if ms_nr_format
          ms_nr_format = ms_nr_format.to_s.downcase
          return nil if ms_nr_format == 'general' or ms_nr_format == ""
          translation_table = {
            'yyyy'=>'%Y',
            'dddd'=>'%A',
            'mmmm'=>'%B',
            'ddd'=>'%a',
            'mmm'=>'%b',
            'yy'=>'%y',
            'dd'=>'%d',
            'mm'=>'%m',
            'y'=>'%y',
            '%%y'=>'%y',
            'd'=>'%e',
            '%%e'=>'%d',
            'm'=>'%m',
            '%%m'=>'%m',
            ';@'=>'',
            '\\'=>''
          }
          translation_table.each{|k,v| ms_nr_format.gsub!(k,v) }
          ms_nr_format
        end
      end

      # Convert numFmtId to msmarkup
      # @param [String, Integer] num_fmt_id numFmtId
      # @return [String] number format (excel markup)
      def num_fmt_id_to_ms_formatting num_fmt_id
        # from: https://stackoverflow.com/questions/4730152/what-indicates-an-office-open-xml-cell-contains-a-date-time-value
        {'0'=>nil, '1'=>'0', '2'=>'0.00', '3'=>'#,##0', '4'=>'#,##0.00',
          '9'=>'0%', '10'=>'0.00%', '11'=>'0.00E+00', '12'=>'# ?/?',
          '13'=>'# ??/??', '14'=>'mm-dd-yy', '15'=>'d-mmm-yy', '16'=>'d-mmm',
          '17'=>'mmm-yy', '18'=>'h:mm AM/PM', '19'=>'h:mm:ss AM/PM',
          '20'=>'h:mm', '21'=>'h:mm:ss', '22'=>'m/d/yy h:mm',
          '37'=>'#,##0 ;(#,##0)', '38'=>'#,##0 ;[Red](#,##0)',
          '39'=>'#,##0.00;(#,##0.00)', '40'=>'#,##0.00;[Red](#,##0.00)',
          '44'=>'_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
          '45'=>'mm:ss', '46'=>'[h]:mm:ss', '47'=>'mmss.0', '48'=>'##0.0E+0',
          '49'=>'@', '27'=>'[$-404]e/m/d', '30'=>'m/d/yy', '36'=>'[$-404]e/m/d',
          '50'=>'[$-404]e/m/d', '57'=>'[$-404]e/m/d', '59'=>'t0', '60'=>'t0.00',
          '61'=>'t#,##0', '62'=>'t#,##0.00', '67'=>'t0%', '68'=>'t0.00%',
          '69'=>'t# ?/?', '70' => 't# ??/??'}[num_fmt_id.to_s]
      end








      # Attempt to convert html-hex color value to xls color number
      #
      # @param [String] hex color
      # @return [String] xls color
      def html_color_to_xls_color hex
        Workbook::Readers::XlsShared::XLS_COLORS.each do |k,v|
          return k if (v == hex or (hex and hex != "" and k == hex.to_sym))
        end
        return nil
      end

      # Converts standard (ruby/C++/unix/...) strftime formatting to MS's formatting
      #
      # @param [String, nil] numberformat (nil returns nil)
      # @return [String, nil]
      def strftime_to_ms_format numberformat
        return nil if numberformat.nil?
        return numberformat.gsub('%Y','yyyy').gsub('%A','dddd').gsub('%B','mmmm').gsub('%a','ddd').gsub('%b','mmm').gsub('%y','yy').gsub('%d','dd').gsub('%m','mm').gsub('%y','y').gsub('%y','%%y').gsub('%e','d')
      end

      def xls_number_to_time number, base_date = DateTime.new(1899,12,30)
        base_date + number.to_f
      end

      def xls_number_to_date number, base_date = Date.new(1899,12,30)
        base_date + number.to_i
      end

      XLS_COLORS = {:xls_color_1=>'#000000',
        :xls_color_2=>'#FFFFFF',
        :xls_color_3=>'#FF0000',
        :xls_color_4=>'#00FF00',
        :xls_color_5=>'#0000FF',
        :xls_color_6=>'#FFFF00',
        :xls_color_7=>'#FF00FF',
        :xls_color_8=>'#00FFFF',
        :xls_color_9=>'#800000',
        :xls_color_10=>'#008000',
        :xls_color_11=>'#000080',
        :xls_color_12=>'#808000',
        :xls_color_13=>'#800080',
        :xls_color_14=>'#008080',
        :xls_color_15=>'#C0C0C0',
        :xls_color_16=>'#808080',
        :xls_color_17=>'#9999FF',
        :xls_color_18=>'#993366',
        :xls_color_19=>'#FFFFCC',
        :xls_color_20=>'#CCFFFF',
        :xls_color_21=>'#660066',
        :xls_color_22=>'#FF8080',
        :xls_color_23=>'#0066CC',
        :xls_color_24=>'#CCCCFF',
        :xls_color_25=>'#000080',
        :xls_color_26=>'#FF00FF',
        :xls_color_27=>'#FFFF00',
        :xls_color_28=>'#00FFFF',
        :xls_color_29=>'#800080',
        :xls_color_30=>'#800000',
        :xls_color_31=>'#008080',
        :xls_color_32=>'#0000FF',
        :xls_color_33=>'#00CCFF',
        :xls_color_34=>'#CCFFFF',
        :xls_color_35=>'#CCFFCC',
        :xls_color_36=>'#FFFF99',
        :xls_color_37=>'#99CCFF',
        :xls_color_38=>'#FF99CC',
        :xls_color_39=>'#CC99FF',
        :xls_color_40=>'#FFCC99',
        :xls_color_41=>'#3366FF',
        :xls_color_42=>'#33CCCC',
        :xls_color_43=>'#99CC00',
        :xls_color_44=>'#FFCC00',
        :xls_color_45=>'#FF9900',
        :xls_color_46=>'#FF6600',
        :xls_color_47=>'#666699',
        :xls_color_48=>'#969696',
        :xls_color_49=>'#003366',
        :xls_color_50=>'#339966',
        :xls_color_51=>'#003300',
        :xls_color_52=>'#333300',
        :xls_color_53=>'#993300',
        :xls_color_54=>'#993366',
        :xls_color_55=>'#333399',
        :xls_color_56=>'#333333',
        :black=>'#000000',
        :white=>'#FFFFFF',
        :red=>'#FF0000',
        :green=>'#00FF00',
        :blue=>'#0000FF',
        :yellow=>'#FFFF00',
        :magenta=>'#FF00FF',
        :cyan=>'#00FFFF',
        :border=>'#FFFFFF',
        :text=>'#000000',
        :lime=>'#00f94c'
      }
    end
  end
end

