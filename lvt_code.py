import pandas as pd
import numpy as np
import math
import time
from datetime import datetime
from scipy.stats import norm
from fpdf import *
import os
import sys
fpdf.set_global("FPDF_CACHE_MODE", 1)

def read_help_file(file_name):
    
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    path_to_help = os.path.abspath(os.path.join(bundle_dir,file_name))
    return path_to_help
  
class compute_score:
    def __init__(self, user_data, xo, alpha=0.05, rtt=0.95, ci='report'):
        self.user_data = user_data
        self.xo = xo
        self.alpha = alpha
        self.rtt = rtt
        self.ci = ci
        self.sd_data = user_data.std()
        self.mean_data = user_data.mean()
        self.se = self.sd_data * math.sqrt(1 - rtt)
        self.k = norm.ppf(1 - alpha/2) * self.se
        
    def perc_rank(self):
        user_pr = round(len(self.user_data[self.user_data <= self.xo]) / len(self.user_data) * 100)
        lower_pr = round(len(self.user_data[self.user_data <= (self.xo - self.k)]) / len(self.user_data) * 100)
        upper_pr = round(len(self.user_data[self.user_data <= (self.xo + self.k)]) / len(self.user_data) * 100)
        pr = f'{user_pr} ({lower_pr}-{upper_pr})'
        if self.ci == 'main':
            return user_pr
        elif self.ci == 'lower':
            return lower_pr
        elif self.ci == 'report':
            return pr
    # perc_rank_after = perc.rank(ci='main')
    
    def interval_rt(self):
        lower_b = self.xo - self.k
        upper_b = self.xo + self.k
        user_t = round((self.xo - self.mean_data)/ self.sd_data * 10 + 50)
        lower_t = round((lower_b - self.mean_data) / self.sd_data * 10 + 50)
        upper_t = round((upper_b - self.mean_data) / self.sd_data * 10 + 50)
        t = f'{user_t} ({lower_t}-{upper_t})'
        if self.ci == 'main':
            return user_t
        elif self.ci == 'lower':
            return lower_t
        elif self.ci == 'report':
            return t
    # interval_rt_after = perc.rank(ci='main')

def set_path(open_path, save_path):
  user_path = open_path

  sample_path = read_help_file("CSDL-with-CI-T-PR.xlsx")
  normal_font = read_help_file("times.ttf")
  bold_font = read_help_file("SVN-Times_New_Roman_Bold.ttf")
  italic_font = read_help_file("SVN-Times_New_Roman_Italic.ttf")
  bold_italic_font = read_help_file("SVN-Times_New_Roman_Bold_Italic.ttf")


  sample_data = pd.read_excel(sample_path, sheet_name='LVT')
  sample_data = sample_data[['S', 'MDR', 'R', 'SHBA', 'BT']]
  user_data = pd.read_csv(user_path, sep='_', header=None)




  df_info = user_data.iloc[0:6, :-1].stack().tolist()
  col = ['PCODE', 'SEX', 'AGE', 'AGEMON', 'EDLEV', 'ADDINF1', 'KBZ', 'FORM', 'VERSION', 'SERIAL', 'SOURCE', 'AWCODE', 
        'DATE', 'TIME1', 'TIME2', 'IDNO', 'DEVICE', 'CANCEL', 'MERGE', '?', 'R', 'MRTR', 'MDR', 'MRTF', 'MDF', 'S', 
        'S1', 'BT', 'SHBA']
  df_data = user_data.iloc[6:,:-2].stack().tolist()
  info_df = pd.DataFrame([df_info], columns=col)
  info_df.replace(r'^\s*$', np.nan, regex=True, inplace=True)









  TEST_VARIABLE = ['Điểm', 'Kết quả bổ sung:', 'Thời gian trung vị cho các câu trả lời chính xác (giây)',
                        'Thời gian trung vị cho các câu trả lời không chính xác (giây)',
                        'Số câu trả lời chính xác',
                        'Số lượng hình ảnh đã xem',
                        'Thời gian thực hiện']
  WORKING_TIME = time.strftime('%M:%S', time.gmtime(int(info_df.BT)))
  MEDIAN_INCORRECT = '-- (1)' if math.isnan(info_df.MDF) else int(info_df.MDF)
  RAW_SCORE = [int(info_df.loc[0, 'S']), '  ', 
              float(info_df.MDR), MEDIAN_INCORRECT, int(info_df.R), int(info_df.SHBA), str(WORKING_TIME) + ' (2)']
  PR = [compute_score(sample_data.iloc[:, 0], RAW_SCORE[0]).perc_rank(), '  ',
        compute_score(sample_data.MDR, float(info_df.MDR), ci='main').perc_rank(), '', '', '', '   ']
  T_SCORE = [compute_score(sample_data.iloc[:, 0], RAW_SCORE[0]).interval_rt(), '  ',
            compute_score(sample_data.MDR, float(info_df.MDR), ci='main').interval_rt(), '', '', '', '   ']
  FINAL_RESULT = pd.DataFrame(
      {'Biến thử nghiệm': TEST_VARIABLE,
      'Số liệu gốc': RAW_SCORE,
      'PR': PR,
      'T': T_SCORE}
      )


  ITEM = list(range(1, 19))
  ANSWER = [int(x) for x in df_data[: 18]]
  PICTURE_VIEWING = [int(x) for x in df_data[80: 98]]
  WORKING_TIME_PRO = [round(float(x)/100, 2) for x in df_data[40: 58]]
  VIEWING_TIME = [round(float(x)/100, 2) for x in df_data[60: 78]]
  RESULT = [int(x) for x in df_data[20: 38]]
  for i in range(3):
    if VIEWING_TIME[i] > 3:
      VIEWING_TIME[i] = '{}!'.format(VIEWING_TIME[i]) 
  for i in range(3, 18):
    if VIEWING_TIME[i] > 4:
      VIEWING_TIME[i] = str(VIEWING_TIME[i]) + '!'
  for i in range(18):
    if RESULT[i] == 1:
      ANSWER[i] = str(ANSWER[i]) + '+'
    else:
      ANSWER[i] = str(ANSWER[i]) + '-'
  protocol = pd.DataFrame(
      {'Mẫu': ITEM,
      'Trả lời': ANSWER,
      'Lượt xem hình ảnh': PICTURE_VIEWING,
      'Thời gian thực hiện': WORKING_TIME_PRO,
      'Thời gian xem': VIEWING_TIME}
  )




  sex = 'nam, ' if int(info_df.loc[0, 'SEX']) == 1 else 'nữ, '
  user_info = f'Năm sinh ??, {sex}' +  str(int(info_df.loc[0, 'AGE'])) + '; ?? năm, Trình độ học vấn bậc ' + str(int(info_df.loc[0, 'EDLEV']))
  date = info_df.loc[0, 'DATE']
  time1 = info_df.loc[0, 'TIME1']
  time2 = info_df.loc[0, 'TIME2']
  duration = ((datetime.strptime(time2, '%H:%M') - datetime.strptime(time1, '%H:%M')).total_seconds())/60
  test_administration = f'Quản lý thử nghiệm: {date} - {time1}...{time2}, Kéo dài: {int(duration)} phút.'




  def output_df_to_pdf(pdf, df):
      # A cell is a rectangular area, possibly framed, which contains some text
      # Set the width and height of cell
      table_cell_width = [116, 19, 18, 18]
      table_cell_height = 5
      # Select a font as Arial, bold, 8
      pdf.set_font('times', '', 11)
      # Loop over to print column names
      cols = df.columns
      i = 0
      for col in cols:
        align = 'C' if col in ['PR', 'T'] else ''
        pdf.cell(table_cell_width[i], table_cell_height, col, border='BT', align=align)
        i += 1
      
      # Line break
      pdf.ln(table_cell_height)
      # Select a font as Arial, regular, 10
      pdf.set_font('times', '', 11)
      # Loop over to print each data in the table
      for row in df.itertuples():
        i = 0
        for col in ['_1', '_2', 'PR', 'T']:
          align = '' if col == '_1' else 'C'
          value = str(getattr(row, col))
          if value == 'Kết quả bổ sung:':
            pdf.set_font('', 'B', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='T')
          elif value == '  ':
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='T')
          elif value in ['Thời gian thực hiện', '   ', str(WORKING_TIME) + ' (2)']:
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='B')
          else:
            pdf.set_font('', '', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='')
          i += 1
        pdf.ln(table_cell_height)




  def output_df_to_pdf_protocol(pdf, df):
      # A cell is a rectangular area, possibly framed, which contains some text
      # Set the width and height of cell
      table_cell_width = [9, 40, 40, 40, 40]
      table_cell_height = 5
      # Select a font as Times, bold, 11
      pdf.set_font('times', 'B', 11)
      # Loop over to print column names
      cols = df.columns
      i = 0
      for col in cols:
        pdf.cell(table_cell_width[i], table_cell_height, col, border=1, align='C')
        i += 1
      
      # Line break
      pdf.ln(table_cell_height)
      # Select a font as Times, regular, 11
      pdf.set_font('times', '', 11)
      # Loop over to print each data in the table
      for row in df.itertuples():
        i = 0
        for col in ['Mẫu', '_2', '_3', '_4', '_5']:
          value = str(getattr(row, col))
          pdf.cell(table_cell_width[i], table_cell_height, value, align='C', border=1)
          i += 1
        pdf.ln(table_cell_height)



  pdf = FPDF()
  pdf.set_left_margin(20)
  pdf.set_top_margin(15)
  pdf.set_right_margin(20)
  pdf.add_font('times', '', normal_font, uni=True);
  pdf.add_font('times', 'B', bold_font, uni=True)
  pdf.add_font('times', 'I', italic_font, uni=True)
  pdf.add_font('times', 'BI', bold_italic_font, uni=True)
  pdf.add_page()
  pdf.set_font('times', 'B', 14)
  pdf.multi_cell(0, 5, info_df.loc[0, 'PCODE'])
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, user_info, border='B')
  pdf.ln(10)
  pdf.set_font('', 'B', 14)
  pdf.cell(0, 5, 'Kiểm tra nhận thức hình ảnh (Visual Pursuit Test - LVT)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Kiểm tra nhận thức trực quan để đánh giá nhận thức có mục tiêu tập trung')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Mẫu thử nghiệm S3 - Mẫu sàng lọc (18 mẫu)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, test_administration, border='B')
  pdf.ln(10)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Kết quả thử nghiệm - Mẫu chuẩn:')
  pdf.ln(5)
  output_df_to_pdf(pdf, FINAL_RESULT)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nhận xét: '), 4, 'Nhận xét: ')
  pdf.set_font('', '', 9)
  pdf.cell(0, 4, 'Tỷ lệ phân cấp (PR) và điểm T là kết quả của so sánh với toàn bộ mẫu so sánh "Người trưởng thành". Khoảng tin cậy được')
  pdf.ln(4)
  pdf.cell(0, 4, 'đưa ra trong ngoặc đơn bên cạnh các điểm so sánh có sai số là 5%.')
  pdf.ln(4)
  pdf.cell(0, 4, '(1) Không thể tính số liệu thô, vì không có mục nào được trả lời đúng')
  pdf.ln(4)
  pdf.cell(0, 4, '(2) Thời gian thực hiện tính bằng phút:giây')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Nhận xét và giải thích về các biến thử nghiệm:')
  pdf.ln(8)
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Số điểm:', border='T')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Số mẫu được giải đúng trong giới hạn thời gian đã định. Biến này tính đến cả tốc độ và độ chính xác của việc thực hiện. Điểm cao cho thấy nhận thức của người trả lời vừa nhanh vừa chính xác khi có được cái nhìn tổng quan.',
                border='B')
  pdf.set_font('', 'I', 9)
  pdf.multi_cell(0, 4, 'Lưu ý:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, '- Các nhận xét và giải thích được đề cập ở trên về các biến thử nghiệm có thể được bật hoặc tắt tại tab "Cài đặt Mở rộng" của Hệ thống Kiểm tra Vienna.')
  pdf.multi_cell(0, 4, '- Bạn có thể tìm thấy mô tả chi tiết của tất cả các biến thử nghiệm và toàn bộ các ghi chú giải thích trong sổ tay thử nghiệm kỹ thuật số (Sổ tay thử nghiệm này có thể được hiển thị và in ra thông qua giao diện người dùng của Hệ thống Thử nghiệm Vienna).')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Giao thức thử nghiệm:')
  pdf.ln(6)
  output_df_to_pdf_protocol(pdf, protocol)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nhận xét: '), 4, 'Nhận xét: ')
  pdf.set_font('', '', 9)
  pdf.cell(0, 4, 'Answer = câu trả lời được chọn (1… 9); + = chính xác, - = không chính xác; Thời gian thực hiện = thời gian trình bày bức')
  pdf.ln(4)
  pdf.multi_cell(0, 4, 'tranh đầu tiên cho đến khi có câu trả lời; Thời gian xem = khoảng thời gian hình ảnh được xem tổng thể (thời gian tính bằng giây); ! = Thời gian xem vượt quá giới hạn thời gian đã hạn định; - = Mẫu không được trình bày.')
  pdf.output(save_path)
  


