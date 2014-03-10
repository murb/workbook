# -*- encoding : utf-8 -*-
module Workbook
  module Readers
    module XlsShared
      def ms_formatting_to_strftime ms_nr_format
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

