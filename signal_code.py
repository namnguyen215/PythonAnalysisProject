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

  sample_data = pd.read_excel(sample_path, sheet_name='SIGNAL')
  sample_data = sample_data[['SUMRV', 'MDDT']]
  user_data = pd.read_csv(user_path, sep='_', header=None)

  df_info = user_data.iloc[0:6, :-1].stack().tolist()
  col = ['PCODE', 'SEX', 'AGE', 'AGEMON', 'EDLEV', 'ADDINF1', 'KBZ', 'FORM', 'VERSION', 'SERIAL', 'SOURCE', 'AWCODE',
        'DATE', 'TIME1', 'TIME2', 'IDNO', 'DEVICE', 'CANCEL', 'MERGE', '?', 'SUMRV', 'SUMR', 'SUMF', 'SUMV', 'SUMA',
        'TOTAL', 'MDT', 'MDDT', 'SDT', 'MWRV', 'MWR', 'MWF', 'MWV', 'MWA']
  df_data = user_data.iloc[6:,:-2].stack().tolist()
  info_df = pd.DataFrame([df_info], columns=col)



  TEST_VARIABLE = ['S??? ch??nh x??c v?? b??? tr???', 'Th???i gian ph??t hi???n trung v??? (gi??y)', 'S??? kh??ng ch??nh x??c']
  RAW_SCORE = [int(info_df.loc[0, 'SUMRV']), float(info_df.loc[0, 'MDDT']),
              int(info_df.loc[0, 'SUMF'])]
  PR = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).perc_rank() for i in range(2)]
  PR.append('')
  T_SCORE = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).interval_rt() for i in range(2)]
  T_SCORE.append('')
  FINAL_RESULT = pd.DataFrame(
      {'Bi???n th??? nghi???m': TEST_VARIABLE,
      'S??? li???u g???c': [str(i) for i in RAW_SCORE],
      'PR': PR,
      'T': T_SCORE}
      )

  plt.rcParams.update(plt.rcParamsDefault)
  plt.rcParams["figure.figsize"] = (10, 2)
  TEST_VARIABLE = ['S??? ch??nh x??c v?? b??? tr???', 'Th???i gian ph??t hi???n trung v??? (gi??y)']
  RAW_SCORE = [int(info_df.loc[0, 'SUMRV']), float(info_df.loc[0, 'MDDT'])]
  x_plot = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() for i in range(2)]))
  y_plot = list(reversed(TEST_VARIABLE[:2]))
  x_error = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() - compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='lower').interval_rt() for i in range(2)]))
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
  # plt.show()

  PARTIAL_INTERVAL = list(range(1, 21))
  CORRECT_AND_DELAYED = [int(x) for x in df_data[:20]]
  OMITTED = [int(x) for x in df_data[60:80]]
  INCORRECT = [int(x) for x in df_data[40:60]]
  MEAN_DETECTION_TIME = []
  for x in df_data[100:120]:
    try:
      MEAN_DETECTION_TIME.append(float(x))
    except:
      MEAN_DETECTION_TIME.append('--')
  protocol = pd.DataFrame(
      {'Kho???ng ngh??? \n ': PARTIAL_INTERVAL,
      'Ch??nh x??c v?? b??? tr??? \n ': CORRECT_AND_DELAYED,
      'B??? b??? s??t \n ': OMITTED,
      'Kh??ng ch??nh x??c \n ': INCORRECT,
      'Th???i gian ph??t hi???n trung b??nh': MEAN_DETECTION_TIME}
  )

  sex = 'nam, ' if int(info_df.loc[0, 'SEX']) == 1 else 'n???, '
  user_info = f'N??m sinh ??, {sex}' +  str(int(info_df.loc[0, 'AGE'])) + '; ?? n??m, Tr??nh ????? h???c v???n b???c ' + str(int(info_df.loc[0, 'EDLEV']))
  date = info_df.loc[0, 'DATE']
  time1 = info_df.loc[0, 'TIME1']
  time2 = info_df.loc[0, 'TIME2']
  duration = ((datetime.strptime(time2, '%H:%M') - datetime.strptime(time1, '%H:%M')).total_seconds())/60
  test_administration = f'Qu???n l?? th??? nghi???m: {date} - {time1}...{time2}, K??o d??i: {int(duration)} ph??t.'

  MEDIAN_DETECTION_TIME = []
  for x in df_data[120:]:
    try:
      MEDIAN_DETECTION_TIME.append(float(x))
    except ValueError:
      MEDIAN_DETECTION_TIME.append(np.nan)

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
          pdf.set_font('', '', 11)
          pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='')
          i += 1
        pdf.ln(table_cell_height)

  def output_df_to_pdf_protocol(pdf, df):
      # A cell is a rectangular area, possibly framed, which contains some text
      # Set the width and height of cell
      table_cell_width = [0, 34, 34, 34, 34, 34]
      table_cell_height = 5
      # Select a font as Times, bold, 11
      pdf.set_font('times', 'B', 11)
      # Loop over to print column names
      cols = df.columns
      y_before = pdf.get_y()
      next_x = 0
      i = 1
      for col in cols:
        pdf.multi_cell(table_cell_width[i], table_cell_height, col, border=1, align='C')
        next_x += table_cell_width[i]
        pdf.set_xy(next_x + pdf.l_margin, y_before)
        i += 1
      
      # Line break
      pdf.ln(table_cell_height*2)
      # Select a font as Times, regular, 11
      pdf.set_font('times', '', 11)
      # Loop over to print each data in the table
      for row in df.itertuples():
        i = 1
        for col in ['_1', '_2', '_3', '_4', '_5']:
          value = str(getattr(row, col))
          pdf.cell(table_cell_width[i], table_cell_height, value, align='C', border=1)
          i += 1
        pdf.ln(table_cell_height)

  idx = np.isfinite(np.arange(1, 21, 1)) & np.isfinite(MEDIAN_DETECTION_TIME)
  b, a = np.polyfit(np.arange(1, 21, 1)[idx], np.array(MEDIAN_DETECTION_TIME)[idx], deg=1)
  plt.rcParams["figure.figsize"] = (17, 13)
  plt.rcParams['axes.spines.right'] = False
  plt.rcParams['axes.spines.top'] = False
  fig, ax = plt.subplots(3, sharex=True)
  # First plot
  ax[0].plot(1, 0, ">k", transform=ax[0].get_yaxis_transform(), clip_on=False)
  ax[0].plot(-0.3, 1, "^k", transform=ax[0].get_xaxis_transform(), clip_on=False)
  ax[0].plot(MEDIAN_DETECTION_TIME, color='b', marker='o')
  ax[0].plot(a + b * np.arange(0, 20, 1), color='r')
  ax[0].set_xlim(-0.3, 19.5)
  ax[0].set_ylim(0, 2.7)
  ax[0].set_yticklabels(['0.0', '0.5', '1.0', '1.5', '2.0', '2.5'])
  ax[0].set_xticks(np.arange(0, 19.5, 1))
  ax[0].set_title('Th???i gian ph???n ???ng trung v??? (gi??y.)', loc='left', fontsize=11)
  # Second plot
  ax[1].plot(1, 0, ">k", transform=ax[1].get_yaxis_transform(), clip_on=False)
  ax[1].plot(-0.3, 1, "^k", transform=ax[1].get_xaxis_transform(), clip_on=False)
  ax[1].plot(np.arange(0, 20, 1), CORRECT_AND_DELAYED, marker='o', color='b')
  ax[1].set_ylim(0, 3.3)
  ax[1].set_yticks(np.arange(0, 3.3, 1))
  ax[1].set_title('S??? ????ng v?? b??? tr???', loc='left', fontsize=11)
  # Third plot
  ax[2].plot(1, 0, ">k", transform=ax[2].get_yaxis_transform(), clip_on=False)
  ax[2].plot(-0.3, 1, "^k", transform=ax[2].get_xaxis_transform(), clip_on=False)
  ax[2].bar(np.arange(0, 20, 1), INCORRECT, width=0.1, color='b')
  ax[2].set_ylim(0, 1.2)
  ax[2].set_yticks(np.arange(0, 1.2, 1))
  ax[2].set_title('Kho???n kh??ng ????ng', loc='left', fontsize=11)
  ax[2].set_xlabel('Kho???ng m???t ph???n', loc='right', fontsize=11)
  ax[2].set_xticklabels([str(x) for x in range(1, 21)])
  test_protocol_temp_file = read_help_file("test_protocol.png")
  plt.savefig(test_protocol_temp_file, format="png", bbox_inches='tight', dpi=200)
  # plt.show()

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
  pdf.cell(0, 5, 'Th??? nghi???m ph??t hi???n t??n hi???u (Signal Detection - SIGNAL)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Th??? nghi???m ????? ????nh gi?? hi???u su???t ch?? ?? c?? ch???n l???c d??i h???n')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'M???u th??? nghi???m S1 - Ti??u chu???n (t??n hi???u tr???ng tr??n n???n ??en)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.multi_cell(0, 5, 'Th??? nghi???m k??o d??i 750 gi??y (20 kho???ng ngh???, m???i kho???ng ngh??? k??o d??i 37.5 gi??y); Ba t??n hi???u quan tr???ng tr??n m???i kho???ng ngh???, m???i t??n hi???u c?? th???i l?????ng t??n hi???u l?? 3,75 gi??y. C??c ph???n ???ng trong v??ng 0,5 gi??y sau khi k???t th??c m???t t??n hi???u quan tr???ng ???????c l??u tr??? d?????i d???ng b??? tr???.')
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
  pdf.cell(0, 5, 'H??? s?? - M???u chu???n:')
  pdf.ln(5)
  pdf.image(test_tmp_path, w=170)
  pdf.add_page()
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, '????? th??? ti???n ?????:')
  pdf.ln(5)
  pdf.image(test_protocol_temp_file, w=170)
  pdf.ln(5)
  pdf.cell(0, 5, 'Giao th???c th??? nghi???m:')
  pdf.ln(5)
  output_df_to_pdf_protocol(pdf, protocol)
  pdf.output(output_path)
  os.remove(test_tmp_path)
  os.remove(test_protocol_temp_file)