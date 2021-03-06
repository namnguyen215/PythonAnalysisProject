# coding=utf8
import pandas as pd
import numpy as np
import math
from datetime import datetime
from scipy.stats import norm
from fpdf import *
import matplotlib.pyplot as plt
import os
import sys
fpdf.set_global("FPDF_CACHE_MODE", 1)

def read_help_file(file_name):
    
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    path_to_help = os.path.abspath(os.path.join(bundle_dir,file_name))
    return path_to_help

sample_path = read_help_file("CSDL-with-CI-T-PR.xlsx")
normal_font = read_help_file("times.ttf")
bold_font = read_help_file("SVN-Times_New_Roman_Bold.ttf")
italic_font = read_help_file("SVN-Times_New_Roman_Italic.ttf")
bold_italic_font = read_help_file("SVN-Times_New_Roman_Bold_Italic.ttf")

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
  output_path = save_path
  sample_data = pd.read_excel(sample_path, sheet_name='DT')
  sample_data = sample_data[['ZV', 'F', 'A']]
  user_data = pd.read_csv(user_path, sep='_', header=None)

  df_info = user_data.iloc[0:4, :-1].stack().tolist()
  col = ['PCODE', 'SEX', 'AGE', 'AGEMON', 'EDLEV', 'ADDINF1', 'KBZ', 'FORM', 'VERSION', 'SERIAL', 'SOURCE', 'AWCODE',
        'DATE', 'TIME1', 'TIME2', 'IDNO', 'DEVICE', 'CANCEL', 'MERGE', '?', 'ZV', 'V', 'F', 'A', 'MDRT', 'S', 'R', 'ADELAY']
  df_data = user_data.iloc[4:,:-2].stack().tolist()
  info_df = pd.DataFrame([df_info], columns=col)



  TEST_VARIABLE = ['Ch??? ????? th??ch ???ng k???t qu??? t???ng th??? (th???i l?????ng th??? nghi???m: 4 ph??t)', 'Ch??nh x??c', 'Kh??ng ch??nh x??c',
                  'B??? b??? qua', 'Th???i gian ph???n ???ng trung v???', 'S??? l?????ng k??ch th??ch', 'C??c ph???n ???ng']
  RAW_SCORE = [int(info_df.loc[0, 'ZV']), int(info_df.loc[0, 'F']), int(info_df.loc[0, 'A']), 
              float(info_df.loc[0, 'MDRT']), int(info_df.loc[0, 'S']), int(info_df.loc[0, 'R'])]
  PR = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).perc_rank() for i in range(3)]
  PR.insert(0, '  ')
  PR.extend([''] * 3)
  T_SCORE = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).interval_rt() for i in range(3)]
  T_SCORE.insert(0, '  ')
  T_SCORE.extend([''] * 3)
  RAW_SCORE.insert(0, '  ')
  FINAL_RESULT = pd.DataFrame(
      {'Bi???n th??? nghi???m': TEST_VARIABLE,
      'S??? li???u g???c': [str(i) for i in RAW_SCORE],
      'PR': PR,
      'T': T_SCORE}
      )

  plt.rcParams.update(plt.rcParamsDefault)
  plt.rcParams["figure.figsize"] = (8.5, 2)
  TEST_VARIABLE = ['Ch??nh x??c', 'Kh??ng ch??nh x??c', 'B??? b??? qua']
  RAW_SCORE = [int(info_df.loc[0, 'ZV']), int(info_df.loc[0, 'F']), int(info_df.loc[0, 'A'])]
  x_plot = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() for i in range(3)]))
  y_plot = list(reversed(TEST_VARIABLE[:3]))
  x_error = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() - compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='lower').interval_rt() for i in range(3)]))
  fig, ax1 = plt.subplots()
  ax2 = ax1.twiny()
  ax1.plot(x_plot, y_plot, marker='o', color='b')
  ax2.plot(x_plot, y_plot, marker='o', color='b')
  ax1.errorbar(x_plot, y_plot, xerr=x_error, color='b', capsize=4)
  ax1.set_xticks(list(range(20, 90, 10)))
  ax2.set_xticks(list(range(20, 90, 10)))
  ax1.set_xticklabels(['0.1', '2.3', '15.9', '50.0', '84.1', '97.7', '99.9'], fontweight='bold')
  ax2.set_xticklabels(['20', '30', '40', '50', '60', '70', '80'], fontweight='bold')
  ax1.set_xlim([10, 90])
  ax2.set_xlim([10, 90])
  ax1.set_xlabel('PR', fontweight='bold')
  ax2.set_xlabel('T', fontweight='bold')
  test_tmp_path = read_help_file("test.png")
  plt.savefig(test_tmp_path, format="png", bbox_inches='tight', dpi=200)

  NBR_STIMULI = int(info_df.loc[0, 'S'])
  MS = [int(x) for x in df_data[:NBR_STIMULI]]
  TIME = []
  for i in range(1620, 1620 + NBR_STIMULI):
    try:
      TIME.append(float(df_data[i]) * 1000)
    except ValueError:
      TIME.append(np.nan)
  INCORRECT = [1 if x == 3 else 0 for x in MS]
  OMITTED = [1 if x == 0 else 0 for x in MS]

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
      # Select a font as times, regular, 11
      pdf.set_font('times', '', 11)
      # Loop over to print each data in the table
      for row in df.itertuples():
        i = 0
        for col in ['_1', '_2', 'PR', 'T']:
          align = '' if col == '_1' else 'C'
          value = str(getattr(row, col))
          if value == 'Ch??? ????? th??ch ???ng k???t qu??? t???ng th??? (th???i l?????ng th??? nghi???m: 4 ph??t)':
            pdf.set_font('', 'B', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='')
          else:
            pdf.set_font('', '', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='')
          i += 1
        pdf.ln(table_cell_height)

  STIMULI_LIMIT = (math.floor(NBR_STIMULI / 20) + 1) * 20
  idx = np.isfinite(np.arange(1, NBR_STIMULI + 1, 1)) & np.isfinite(TIME)
  b, a = np.polyfit(np.arange(1, NBR_STIMULI + 1, 1)[idx], np.array(TIME)[idx], deg=1)
  plt.rcParams["figure.figsize"] = (17, 13)
  plt.rcParams['axes.spines.right'] = False
  plt.rcParams['axes.spines.top'] = False
  fig, ax = plt.subplots(3, sharex=True)
  # First plot
  ax[0].plot(1, 500, ">k", transform=ax[0].get_yaxis_transform(), clip_on=False)
  ax[0].plot(-5, 1, "^k", transform=ax[0].get_xaxis_transform(), clip_on=False)
  ax[0].scatter(np.arange(0, NBR_STIMULI, 1), TIME, color='b')
  ax[0].plot(np.arange(1, NBR_STIMULI+1, 1), a + b * np.arange(1, NBR_STIMULI+1, 1), color='r')
  ax[0].set_xlim(-5, STIMULI_LIMIT + 3)
  ax[0].set_ylim(500, 1550)
  ax[0].set_xticks(np.arange(0, STIMULI_LIMIT + 2, 20))
  ax[0].set_yticks(np.arange(500, 1550, 100))
  ax[0].set_title('Th???i gian ph???n ???ng (t??nh b???ng mili gi??y)', loc='left', fontsize=11)
  # Second plot
  ax[1].plot(1, 0, ">k", transform=ax[1].get_yaxis_transform(), clip_on=False)
  ax[1].plot(-5, 1, "^k", transform=ax[1].get_xaxis_transform(), clip_on=False)
  ax[1].bar(np.arange(1, NBR_STIMULI+1, 1), OMITTED, width=2, color='b')
  ax[1].set_ylim(0, 1.3)
  ax[1].set_yticks(np.arange(0, 1.3, 1))
  ax[1].set_title('B??? b??? qua', loc='left', fontsize=11)
  # Third plot
  ax[2].plot(1, 0, ">k", transform=ax[2].get_yaxis_transform(), clip_on=False)
  ax[2].plot(-5, 1, "^k", transform=ax[2].get_xaxis_transform(), clip_on=False)
  ax[2].bar(np.arange(1, NBR_STIMULI+1, 1), INCORRECT, width=2, color='b')
  ax[2].set_ylim(0, 1.3)
  ax[2].set_yticks(np.arange(0, 1.3, 1))
  ax[2].set_title('Kh??ng ch??nh x??c', loc='left', fontsize=11)
  ax[2].set_xlabel('K??ch th??ch', loc='right', fontsize=11)
  test_protocol_temp_file = read_help_file("test_protocol.png")
  plt.savefig(test_protocol_temp_file, format="png", bbox_inches='tight', dpi=200)

  pdf = FPDF()
  pdf.set_left_margin(20)
  pdf.set_top_margin(15)
  pdf.set_right_margin(20)
  pdf.add_font('times', '', normal_font, uni=True)
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
  pdf.cell(0, 5, 'Th??? nghi???m x??c ?????nh (Determination Test - DT)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Th?? nghi???m ph???n ???ng tr???c nghi???m c?? k??ch th??ch ph???c t???p')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'M???u th??? nghi???m S1 - M???u ng???n c?? tr??nh b??y k??ch th??ch th??ch ???ng (t???t c??? c??c lo???i k??ch th??ch)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Th??? nghi???m k??o d??i 4 ph??t')
  pdf.ln(5)
  pdf.cell(0, 5, test_administration, border='B')
  pdf.ln(10)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'K???t qu??? th??? nghi???m - M???u chu???n:')
  pdf.ln(5)
  output_df_to_pdf(pdf, FINAL_RESULT)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nh???n x??t: '), 4, 'Nh???n x??t: ', border='T')
  pdf.set_font('', '', 9)
  pdf.cell(0, 4, 'T??? l??? ph??n c???p (PR) v?? ??i???m T l?? k???t qu??? c???a so s??nh v???i to??n b??? m???u so s??nh ti??u chu???n. Kho???ng tin c???y ???????c ????a ra', border='T')
  pdf.ln(4)
  pdf.cell(0, 4, 'trong ngo???c ????n b??n c???nh c??c ??i???m so s??nh c?? sai s??? l?? 5%.')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Nh???n x??t v?? gi???i th??ch v??? c??c bi???n th??? nghi???m:')
  pdf.ln(8)
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Ch??nh x??c:', border='T')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Bi???n ch??nh n??y n??u chi ti???t t???ng s??? ph???n ???ng ch??nh x??c ???????c th???c hi???n tr?????c khi b???t ?????u h??nh ?????ng ti???p theo tr??? m???t k??ch th??ch. N?? ??o l?????ng kh??? n??ng ti???p t???c ph???n ???ng nhanh ch??ng v?? th??ch h???p trong chu???i ph???n ???ng c???a ng?????i tr??? l???i ?????, bao g???m c??? khi l??m vi???c g???n ?????n gi???i h???n ch???u ?????ng c??ng th???ng c???a c?? nh??n h???.',
                border='B')
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Kh??ng ch??nh x??c:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Bi???n n??y m?? t??? xu h?????ng nh???m l???n gi???a c??c ph???n ???ng kh??c nhau. Ph???n ???ng kh??ng ch??nh x??c ph??t sinh b???i v?? khi b??? c??ng th???ng, ng?????i tr??? l???i kh??ng th??? b???o v??? ph???n ???ng th??ch h???p kh???i ???nh h?????ng c???a c??c k??ch th??ch kh??ng th??ch h???p.',
                border='B')
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'B??? b??? qua:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Bi???n ph??? n??y cho bi???t li???u c??c ph???n h???i c?? b??? b??? qua d?????i ??p l???c th???i gian hay kh??ng. Nh???ng c?? nh??n c?? ??i???m r???t cao v??? bi???n n??y c?? xu h?????ng kh??ng th??? duy tr?? s??? ch?? ?? c???a h??? khi th???c hi???n c??c nhi???m v??? thu???c lo???i n??y trong t??nh tr???ng c??ng th???ng; ??i???u n??y c?? ngh??a l?? trong nh???ng t??nh hu???ng c??ng th???ng, h??? c?? th??? c?? xu h?????ng b??? cu???c.',
                border='B')
  pdf.set_font('', 'I', 9)
  pdf.multi_cell(0, 4, 'L??u ??:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, '- C??c nh???n x??t v?? gi???i th??ch ???????c ????? c???p ??? tr??n v??? c??c bi???n th??? nghi???m c?? th??? ???????c b???t ho???c t???t t???i tab "C??i ?????t M??? r???ng" c???a H??? th???ng Ki???m tra Vienna.')
  pdf.multi_cell(0, 4, '- B???n c?? th??? t??m th???y m?? t??? chi ti???t c???a t???t c??? c??c bi???n th??? nghi???m v?? to??n b??? c??c ghi ch?? gi???i th??ch trong s??? tay th??? nghi???m k??? thu???t s??? (S??? tay th??? nghi???m n??y c?? th??? ???????c hi???n th??? v?? in ra th??ng qua giao di???n ng?????i d??ng c???a H??? th???ng Th??? nghi???m Vienna).')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'H??? s?? - M???u chu???n:')
  pdf.ln(5)
  pdf.image(test_tmp_path, w=170)
  pdf.add_page()
  pdf.cell(0, 5, '????? th??? ti???n ?????:')
  pdf.ln(5)
  pdf.image(test_protocol_temp_file, w=170)
  pdf.ln(5)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nh???n x??t: '), 4, 'Nh???n x??t: ')
  pdf.set_font('', '', 9)
  pdf.set_text_color(255, 0, 0)
  pdf.cell(pdf.get_string_width('?????? '), 4, '?????? ')
  pdf.set_text_color(0, 0, 0)
  pdf.cell(0, 4, '???????ng h???i quy')
  pdf.output(output_path)
  os.remove(test_tmp_path)
  os.remove(test_protocol_temp_file)
  