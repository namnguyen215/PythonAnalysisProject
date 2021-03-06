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









  TEST_VARIABLE = ['??i???m', 'K???t qu??? b??? sung:', 'Th???i gian trung v??? cho c??c c??u tr??? l???i ch??nh x??c (gi??y)',
                        'Th???i gian trung v??? cho c??c c??u tr??? l???i kh??ng ch??nh x??c (gi??y)',
                        'S??? c??u tr??? l???i ch??nh x??c',
                        'S??? l?????ng h??nh ???nh ???? xem',
                        'Th???i gian th???c hi???n']
  WORKING_TIME = time.strftime('%M:%S', time.gmtime(int(info_df.BT)))
  MEDIAN_INCORRECT = '-- (1)' if math.isnan(info_df.MDF) else int(info_df.MDF)
  RAW_SCORE = [int(info_df.loc[0, 'S']), '  ', 
              float(info_df.MDR), MEDIAN_INCORRECT, int(info_df.R), int(info_df.SHBA), str(WORKING_TIME) + ' (2)']
  PR = [compute_score(sample_data.iloc[:, 0], RAW_SCORE[0]).perc_rank(), '  ',
        compute_score(sample_data.MDR, float(info_df.MDR), ci='main').perc_rank(), '', '', '', '   ']
  T_SCORE = [compute_score(sample_data.iloc[:, 0], RAW_SCORE[0]).interval_rt(), '  ',
            compute_score(sample_data.MDR, float(info_df.MDR), ci='main').interval_rt(), '', '', '', '   ']
  FINAL_RESULT = pd.DataFrame(
      {'Bi???n th??? nghi???m': TEST_VARIABLE,
      'S??? li???u g???c': RAW_SCORE,
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
      {'M???u': ITEM,
      'Tr??? l???i': ANSWER,
      'L?????t xem h??nh ???nh': PICTURE_VIEWING,
      'Th???i gian th???c hi???n': WORKING_TIME_PRO,
      'Th???i gian xem': VIEWING_TIME}
  )




  sex = 'nam, ' if int(info_df.loc[0, 'SEX']) == 1 else 'n???, '
  user_info = f'N??m sinh ??, {sex}' +  str(int(info_df.loc[0, 'AGE'])) + '; ?? n??m, Tr??nh ????? h???c v???n b???c ' + str(int(info_df.loc[0, 'EDLEV']))
  date = info_df.loc[0, 'DATE']
  time1 = info_df.loc[0, 'TIME1']
  time2 = info_df.loc[0, 'TIME2']
  duration = ((datetime.strptime(time2, '%H:%M') - datetime.strptime(time1, '%H:%M')).total_seconds())/60
  test_administration = f'Qu???n l?? th??? nghi???m: {date} - {time1}...{time2}, K??o d??i: {int(duration)} ph??t.'




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
          if value == 'K???t qu??? b??? sung:':
            pdf.set_font('', 'B', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='T')
          elif value == '  ':
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='T')
          elif value in ['Th???i gian th???c hi???n', '   ', str(WORKING_TIME) + ' (2)']:
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
        for col in ['M???u', '_2', '_3', '_4', '_5']:
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
  pdf.cell(0, 5, 'Ki???m tra nh???n th???c h??nh ???nh (Visual Pursuit Test - LVT)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Ki???m tra nh???n th???c tr???c quan ????? ????nh gi?? nh???n th???c c?? m???c ti??u t???p trung')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'M???u th??? nghi???m S3 - M???u s??ng l???c (18 m???u)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, test_administration, border='B')
  pdf.ln(10)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'K???t qu??? th??? nghi???m - M???u chu???n:')
  pdf.ln(5)
  output_df_to_pdf(pdf, FINAL_RESULT)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nh???n x??t: '), 4, 'Nh???n x??t: ')
  pdf.set_font('', '', 9)
  pdf.cell(0, 4, 'T??? l??? ph??n c???p (PR) v?? ??i???m T l?? k???t qu??? c???a so s??nh v???i to??n b??? m???u so s??nh "Ng?????i tr?????ng th??nh". Kho???ng tin c???y ???????c')
  pdf.ln(4)
  pdf.cell(0, 4, '????a ra trong ngo???c ????n b??n c???nh c??c ??i???m so s??nh c?? sai s??? l?? 5%.')
  pdf.ln(4)
  pdf.cell(0, 4, '(1) Kh??ng th??? t??nh s??? li???u th??, v?? kh??ng c?? m???c n??o ???????c tr??? l???i ????ng')
  pdf.ln(4)
  pdf.cell(0, 4, '(2) Th???i gian th???c hi???n t??nh b???ng ph??t:gi??y')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Nh???n x??t v?? gi???i th??ch v??? c??c bi???n th??? nghi???m:')
  pdf.ln(8)
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'S??? ??i???m:', border='T')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'S??? m???u ???????c gi???i ????ng trong gi???i h???n th???i gian ???? ?????nh. Bi???n n??y t??nh ?????n c??? t???c ????? v?? ????? ch??nh x??c c???a vi???c th???c hi???n. ??i???m cao cho th???y nh???n th???c c???a ng?????i tr??? l???i v???a nhanh v???a ch??nh x??c khi c?? ???????c c??i nh??n t???ng quan.',
                border='B')
  pdf.set_font('', 'I', 9)
  pdf.multi_cell(0, 4, 'L??u ??:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, '- C??c nh???n x??t v?? gi???i th??ch ???????c ????? c???p ??? tr??n v??? c??c bi???n th??? nghi???m c?? th??? ???????c b???t ho???c t???t t???i tab "C??i ?????t M??? r???ng" c???a H??? th???ng Ki???m tra Vienna.')
  pdf.multi_cell(0, 4, '- B???n c?? th??? t??m th???y m?? t??? chi ti???t c???a t???t c??? c??c bi???n th??? nghi???m v?? to??n b??? c??c ghi ch?? gi???i th??ch trong s??? tay th??? nghi???m k??? thu???t s??? (S??? tay th??? nghi???m n??y c?? th??? ???????c hi???n th??? v?? in ra th??ng qua giao di???n ng?????i d??ng c???a H??? th???ng Th??? nghi???m Vienna).')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Giao th???c th??? nghi???m:')
  pdf.ln(6)
  output_df_to_pdf_protocol(pdf, protocol)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nh???n x??t: '), 4, 'Nh???n x??t: ')
  pdf.set_font('', '', 9)
  pdf.cell(0, 4, 'Answer = c??u tr??? l???i ???????c ch???n (1??? 9); + = ch??nh x??c, - = kh??ng ch??nh x??c; Th???i gian th???c hi???n = th???i gian tr??nh b??y b???c')
  pdf.ln(4)
  pdf.multi_cell(0, 4, 'tranh ?????u ti??n cho ?????n khi c?? c??u tr??? l???i; Th???i gian xem = kho???ng th???i gian h??nh ???nh ???????c xem t???ng th??? (th???i gian t??nh b???ng gi??y); ! = Th???i gian xem v?????t qu?? gi???i h???n th???i gian ???? h???n ?????nh; - = M???u kh??ng ???????c tr??nh b??y.')
  pdf.output(save_path)
  


