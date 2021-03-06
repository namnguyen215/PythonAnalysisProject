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
  sample_data = pd.read_excel(sample_path, sheet_name='RT')
  sample_data = sample_data[['MRZ', 'MMZ', 'SDRZ', 'SDMZ']]
  user_data = pd.read_csv(user_path, sep='_', header=None)

  df_info = user_data.iloc[0:6, :-1].stack().tolist()
  col = ['PCODE', 'SEX', 'AGE', 'AGEMON', 'EDLEV', 'ADDINF1', 'KBZ', 'FORM', 'VERSION', 'SERIAL', 'SOURCE', 'AWCODE',
        'DATE', 'TIME1', 'TIME2', 'IDNO', 'DEVICE', 'CANCEL', 'MERGE', '?', 'MRZ', 'MMZ', 'SDRZ', 'SDMZ', 'RR', 'NR',
        'UR', 'FR', 'OMRZ', 'OMMZ', 'OSDRZ', 'OSDMZ', 'LMRZ', 'LMMZ', 'RMRZ', 'RMMZ', 'ADELAY']
  df_data = user_data.iloc[6:,:-2].stack().tolist()
  info_df = pd.DataFrame([df_info], columns=col)


  TEST_VARIABLE = ['Th???i gian ph???n ???ng trung b??nh (2)', 'Th???i gian v???n ?????ng trung b??nh (2)', '??o th???i gian ph???n ???ng ph??n t??n (3)',
                  '??o th???i gian v???n ?????ng ph??n t??n (3)', 'K???t qu??? b??? sung', 'Ph???n ???ng ch??nh x??c', 'Kh??ng c?? ph???n ???ng',
                  'Ph???n ???ng kh??ng ho??n ch???nh', 'Ph???n ???ng kh??ng ch??nh x??c']
  RAW_SCORE = [int(info_df.loc[0, 'MRZ']), int(info_df.loc[0, 'MMZ']), int(info_df.loc[0, 'SDRZ']), int(info_df.loc[0, 'SDMZ']), '  ', 
              int(info_df.loc[0, 'RR']), int(info_df.loc[0, 'NR']), int(info_df.loc[0, 'UR']), int(info_df.loc[0, 'FR'])]
  PR = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).perc_rank() for i in range(4)]
  PR.extend(['  '] + [''] * 4)
  T_SCORE = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).interval_rt() for i in range(4)]
  T_SCORE.extend(['  '] + [''] * 4)
  FINAL_RESULT = pd.DataFrame(
      {'Bi???n th??? nghi???m': TEST_VARIABLE,
      'S??? li???u g???c': [str(i) for i in RAW_SCORE],
      'PR': PR,
      'T': T_SCORE}
      )
  plt.rcParams.update(plt.rcParamsDefault)
  plt.rcParams["figure.figsize"] = (8.5, 2)
  TEST_VARIABLE = ['Th???i gian ph???n ???ng trung b??nh', 'Th???i gian v???n ?????ng trung b??nh',
                  '??o th???i gian ph???n ???ng ph??n t??n', '??o th???i gian v???n ?????ng ph??n t??n']
  RAW_SCORE = [int(info_df.loc[0, 'MRZ']), int(info_df.loc[0, 'MMZ']),
              float(info_df.loc[0, 'SDRZ']), float(info_df.loc[0, 'SDMZ'])]
  x_plot = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() for i in range(4)]))
  y_plot = list(reversed(TEST_VARIABLE[:4]))
  x_error = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() - compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='lower').interval_rt() for i in range(4)]))
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

  REACTION_TIME = []
  MOTOR_TIME = []
  for x in df_data[144:192]:
    try:
      REACTION_TIME.append(float(x))
    except ValueError:
      REACTION_TIME.append(np.nan)
  for x in df_data[192:]:
    try:
      MOTOR_TIME.append(float(x))
    except ValueError:
      MOTOR_TIME.append(np.nan)
  STIM_NR = list(range(1, 49))
  STIM_TYPE = ['V??ng + T??ng m??u', '?????', 'V??ng', 'T??ng m??u', '????? + T??ng m??u', 'V??ng', 'V??ng + T??ng m??u', '?????', 'V??ng', 'T??ng m??u', '????? + T??ng m??u', 'V??ng', 'V??ng + T??ng m??u', '????? + T??ng m??u', 'V??ng', 'V??ng + T??ng m??u','?????', 'V??ng + T??ng m??u', 'V??ng + T??ng m??u', '????? + T??ng m??u', 'V??ng', 'V??ng + T??ng m??u', '?????', 'V??ng + T??ng m??u','V??ng + T??ng m??u', '?????', 'V??ng', 'T??ng m??u', '????? + T??ng m??u', 'V??ng', 'V??ng + T??ng m??u', '?????','V??ng', 'T??ng m??u', '????? + T??ng m??u', 'V??ng', 'V??ng + T??ng m??u', '????? + T??ng m??u', 'V??ng', 'V??ng + T??ng m??u','?????', 'V??ng + T??ng m??u', 'V??ng + T??ng m??u', '????? + T??ng m??u', 'V??ng', 'V??ng + T??ng m??u', '?????', 'V??ng + T??ng m??u']
  REQUIRED = ['Yes' if int(x) == 1 else 'No' for x in df_data[48:96]]
  EVALUATION = []
  for x in df_data[96:144]:
    # try:
      if int(x) == 1:
        EVALUATION.append('Ph???n ???ng kh??ng ch??nh x??c')
      elif int(x) == 2:
        EVALUATION.append('Kh??ng ph???n ???ng')
      elif int(x) == 3:
          EVALUATION.append('Ph???n ???ng kh??ng ho??n ch???nh')
      elif int(x) == 4:
        EVALUATION.append('Ph???n ???ng ch??nh x??c')
      else: EVALUATION.append('??????')

  idx_1 = np.isfinite(np.arange(1, 49, 1)) & np.isfinite(REACTION_TIME)
  xticks_label = []
  for i in range(1, 49, 2):
    xticks_label.extend(['', str(i+1)])
  b, a = np.polyfit(np.arange(1, 49, 1)[idx_1], np.array(REACTION_TIME)[idx_1], deg=1)
  b1, a1 = np.polyfit(np.arange(1, 49, 1)[idx_1], np.array(MOTOR_TIME)[idx_1], deg=1)
  plt.rcParams["figure.figsize"] = (17, 20)
  plt.rcParams['axes.spines.right'] = False
  plt.rcParams['axes.spines.top'] = False
  fig, ax = plt.subplots(2, sharex=True)
  # First plot
  ax[0].plot(1, 0, ">k", transform=ax[0].get_yaxis_transform(), clip_on=False)
  ax[0].plot(0, 1, "^k", transform=ax[0].get_xaxis_transform(), clip_on=False)
  ax[0].scatter(np.arange(1, 49, 1)[idx_1], np.array(REACTION_TIME)[idx_1], color='b')
  ax[0].plot(np.arange(1, 49, 1), a + b * np.arange(1, 49, 1), color='r')
  ax[0].set_xlim(0, 48.5)
  ax[0].set_ylim(0, 1050)
  ax[0].set_xticks(np.arange(1, 48.5, 1))
  ax[0].set_yticks(np.arange(0, 1050, 100))
  ax[0].set_title('Th???i gian ph???n ???ng t??nh b???ng mili gi??y', loc='left', fontsize=11)
  # Second plot
  ax[1].plot(1, 0, ">k", transform=ax[1].get_yaxis_transform(), clip_on=False)
  ax[1].plot(0, 1, "^k", transform=ax[1].get_xaxis_transform(), clip_on=False)
  ax[1].scatter(np.arange(1, 49, 1)[idx_1], np.array(MOTOR_TIME)[idx_1], color='b')
  ax[1].plot(np.arange(1, 49, 1), a1 + b1 * np.arange(1, 49, 1), color='r')
  ax[1].set_xlim(0, 48.5)
  ax[1].set_ylim(0, 650)
  ax[1].set_xticks(np.arange(1, 48.5, 1))
  ax[1].set_yticks(np.arange(0, 650, 100))
  ax[1].set_xticklabels(xticks_label)
  ax[1].set_title('Th???i gian v???n ?????ng t??nh b???ng mili gi??y', loc='left', fontsize=11)
  ax[1].set_xlabel('K??ch th??ch', loc='right', fontsize=11)
  test_protocol_temp_file = read_help_file("test_protocol.png")
  plt.savefig(test_protocol_temp_file, format="png", bbox_inches='tight', dpi=200)

  REACTION_TIME = ['??????' if math.isnan(x) else int(x) for x in REACTION_TIME]
  MOTOR_TIME = ['??????' if math.isnan(x) else int(x) for x in MOTOR_TIME]
  for i in range(len(MOTOR_TIME)):
    if EVALUATION[i] == 'Ph???n ???ng kh??ng ch??nh x??c':
      REACTION_TIME[i] = f'({REACTION_TIME[i]})'
      MOTOR_TIME[i] = f'({MOTOR_TIME[i]})'
  protocol = pd.DataFrame(
      {'STT \n ': STIM_NR,
      'Lo???i k??ch th??ch \n ': STIM_TYPE,
      'Y??u c???u \n ': REQUIRED,
      '????nh gi?? \n ': EVALUATION,
      'Th???i gian ph???n ???ng (mili gi??y)': REACTION_TIME,
      'Th???i gian v???n ?????ng (mili gi??y)': MOTOR_TIME}
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
          if value in ['K???t qu??? b??? sung', '  ']:
            pdf.set_font('', 'B', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='T')
          else:
            pdf.set_font('', '', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='')
          i += 1
        pdf.ln(table_cell_height)

  def output_df_to_pdf_protocol(pdf, df):
      # A cell is a rectangular area, possibly framed, which contains some text
      # Set the width and height of cell
      table_cell_width = [0, 12, 31, 25, 46, 32, 32]
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
        for col in ['_1', '_2', '_3', '_4', '_5', '_6']:
          value = str(getattr(row, col))
          pdf.cell(table_cell_width[i], table_cell_height, value, align='C', border=1)
          i += 1
        pdf.ln(table_cell_height)

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
  pdf.cell(0, 5, 'Ki???m tra ph???n ???ng (Reaction Test - RT)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Th??? nghi???m ????? ????nh gi?? th???i gian ph???n ???ng ?????i v???i c??c k??ch th??ch b???ng ??m thanh v?? h??nh ???nh.')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'M???u th??? S3 - M??u v??ng / ??m ph???n ???ng l???a ch???n')
  pdf.ln(5)
  pdf.set_font('', '', 11)
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
  pdf.ln(4)
  pdf.cell(0, 4, '(1) T???t c??? c??c m???c th???i gian t??nh b???ng mili gi??y')
  pdf.ln(4)
  pdf.cell(0, 4, '(2) Th???i gian ph???n ???ng trung b??nh sau khi chu???n h??a Box-Cox')
  pdf.ln(4)
  pdf.cell(0, 4, '(3) ????? l???ch chu???n c???a th???i gian ph???n ???ng sau khi chu???n h??a Box-Cox ')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Nh???n x??t v?? gi???i th??ch v??? c??c bi???n th??? nghi???m:')
  pdf.ln(8)
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Th???i gian ph???n ???ng trung b??nh:', border='T')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Khi s??? d???ng n??t ngh???, th???i gian ph???n ???ng l?? kho???ng th???i gian t??? khi b???t ?????u t??c ?????ng k??ch th??ch li??n quan ?????n th???i ??i???m ng??n tay r???i kh???i n??t ngh???.')
  pdf.multi_cell(0, 4, '??i???m n??y l?? th???i gian ph???n ???ng c???a m???i l???n th??? nghi???m. ??i???m cao cho th???y r???ng so v???i d??n s??? tham chi???u, ng?????i ???????c h???i c?? kh??? n??ng ph???n ???ng v???i c??c k??ch th??ch ho???c ch??m k??ch th??ch c?? li??n quan nhanh tr??n m???c trung b??nh.',
                border='B')
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Th???i gian v???n ?????ng trung b??nh:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Th???i gian v???n ?????ng l?? th???i gian t??? khi ng??n tay r???i kh???i n??t ngh??? ?????n khi ng??n tay ti???p x??c v???i n??t ph???n ???ng khi m???t k??ch th??ch li??n quan xu???t hi???n.')
  pdf.multi_cell(0, 4, '??i???m n??y cung c???p th??ng tin v??? t???c ????? di chuy???n c???a ng?????i tr??? l???i. ??i???m s??? cao cho th???y r???ng so v???i d??n s??? tham chi???u, ng?????i ???????c h???i c?? kh??? n??ng th???c hi???n m???t c??ch nhanh ch??ng qu?? tr??nh h??nh ?????ng ???? l??n k??? ho???ch trong c??c t??nh hu???ng ph???n ???ng tr??n m???c trung b??nh.',
                border='B')
  pdf.set_font('', 'I', 9)
  pdf.multi_cell(0, 4, 'L??u ??:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, '- C??c nh???n x??t v?? gi???i th??ch ???????c ????? c???p ??? tr??n v??? c??c bi???n th??? nghi???m c?? th??? ???????c b???t ho???c t???t t???i tab "C??i ?????t M??? r???ng" c???a H??? th???ng Ki???m tra Vienna.')
  pdf.multi_cell(0, 4, '- B???n c?? th??? t??m th???y m?? t??? chi ti???t c???a t???t c??? c??c bi???n th??? nghi???m v?? to??n b??? c??c ghi ch?? gi???i th??ch trong s??? tay th??? nghi???m k??? thu???t s??? (S??? tay th??? nghi???m n??y c?? th??? ???????c hi???n th??? v?? in ra th??ng qua giao di???n ng?????i d??ng c???a H??? th???ng Th??? nghi???m Vienna).')
  pdf.add_page()
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
  pdf.set_text_color(0, 0, 255)
  pdf.cell(pdf.get_string_width('??? '), 4, '??? ')
  pdf.set_text_color(0, 0, 0)
  pdf.cell(pdf.get_string_width('Th???i gian ph???n ???ng v???i m???i k??ch th??ch;  '), 4, 'Th???i gian ph???n ???ng v???i m???i k??ch th??ch; ')
  pdf.set_text_color(255, 0, 0)
  pdf.cell(pdf.get_string_width('?????? '), 4, '?????? ')
  pdf.set_text_color(0, 0, 0)
  pdf.cell(0, 4, '???????ng h???i quy')
  pdf.add_page()
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Giao th???c th??? nghi???m:')
  pdf.ln(5)
  output_df_to_pdf_protocol(pdf, protocol)
  pdf.output(output_path)
  os.remove(test_tmp_path)
  os.remove(test_protocol_temp_file)

# set_path('C:/Users/NguyenNam/Desktop/Pilot_Analysis/bach cong thanh/RT03.ASC','result.pdf')