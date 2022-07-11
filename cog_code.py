

import pandas as pd
import numpy as np
import math
import time
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

def set_path(open_path, save_path):
  user_path = open_path
  output_path = save_path

  sample_data = pd.read_excel(sample_path, sheet_name='COG')
  sample_data = sample_data[['SUMMR', 'SUMMF', 'MTRA', 'MTFA']]
  user_data = pd.read_csv(user_path, sep='_', header=None)

  df_info = user_data.iloc[0:3, :-1].stack().tolist()
  col = ['PCODE', 'SEX', 'AGE', 'AGEMON', 'EDLEV', 'ADDINF1', 'KBZ', 'FORM', 'VERSION', 'SERIAL', 'SOURCE', 'AWCODE',
        'DATE', 'TIME1', 'TIME2', 'IDNO', 'DEVICE', 'CANCEL', 'MERGE', '?', 'SUMMR', 'SUMMF', 'SUMMA', 'MTRA', 'MTFA',
        'BT']
  df_data = user_data.iloc[3:,:-2].stack().tolist()
  info_df = pd.DataFrame([df_info], columns=col)

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

  TEST_VARIABLE = ['Tổng số "phản ứng chính xác"', 'Tổng số "phản ứng không chính xác"', 'Tổng số "không phản ứng không chính xác" (1)',
                  'Thời gian "phản ứng chính xác" trung bình (giây)', 'Thời gian "phản ứng không chính xác trung bình"']
  RAW_SCORE = [int(info_df.loc[0, 'SUMMR']), int(info_df.loc[0, 'SUMMF']),
              float(info_df.loc[0, 'MTRA']), float(info_df.loc[0, 'MTFA'])]
  PR = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).perc_rank() for i in range(2)]
  PR.extend([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').perc_rank() for i in range(2, 4)])
  PR.insert(2, '  ')
  T_SCORE = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).interval_rt() for i in range(2)]
  T_SCORE.extend([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() for i in range(2, 4)])
  T_SCORE.insert(2, '  ')
  RAW_SCORE.insert(2, int(info_df.loc[0, 'SUMMA']))
  FINAL_RESULT = pd.DataFrame(
      {'Biến thử nghiệm': TEST_VARIABLE,
      'Số liệu gốc': [str(i) for i in RAW_SCORE],
      'PR': PR,
      'T': T_SCORE}
      )

  plt.rcParams.update(plt.rcParamsDefault)
  plt.rcParams["figure.figsize"] = (8.5, 2)
  TEST_VARIABLE = ['Tổng số "phản ứng chính xác"', 'Tổng số "phản ứng không chính xác"',
                  'Thời gian "phản ứng chính xác" trung bình (giây)', 'Thời gian "phản ứng không chính xác trung bình"']
  RAW_SCORE = [int(info_df.loc[0, 'SUMMR']), int(info_df.loc[0, 'SUMMF']),
              float(info_df.loc[0, 'MTRA']), float(info_df.loc[0, 'MTFA'])]
  x_plot = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() for i in range(4)]))
  y_plot = list(reversed(TEST_VARIABLE[:4]))
  x_error = [0, 0] + list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() - compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='lower').interval_rt() for i in range(2)]))
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

  SAMPLE = ['Mẫu', 'Kích thích'] + [''] * 9
  SAMPLE_2 = [''] + list(range(1, 11))
  REACTION_LABEL = np.reshape(['?+' if int(x) == 0 else '1-' if int(x) == 1 else '?-' if int(x) == 10 else '1+' for x in df_data[:200]], (20, 10))
  REACTION_LABEL = pd.DataFrame(REACTION_LABEL)
  # REACTION_LABEL_FINAL = REACTION_LABEL
  for i in range(200, 400):
    try:
      df_data[i] = float(df_data[i])
    except:
      df_data[i] = 1.8
  REACTION_TIME = np.reshape(df_data[200:400], (20, 10))
  REACTION_TIME = pd.DataFrame(REACTION_TIME)
  # REACTION_TIME_FINAL = REACTION_TIME

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
          pdf.set_font('', '', 11)
          pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='')
          i += 1
        pdf.ln(table_cell_height)

  def output_df_to_pdf_protocol(pdf):
    pdf.set_font('times', 'B', 11)
    pdf.multi_cell(19, 10, 'Mẫu', border=1, align='C')
    y_before = pdf.get_y()
    pdf.set_xy(pdf.l_margin + 19, y_before - 10)
    pdf.multi_cell(150, 5, 'Kích thích', border=1, align='C')
    pdf.set_xy(pdf.l_margin + 19, y_before - 5)
    for i in range(1, 11):
      pdf.cell(15, 5, str(i), align='C', border=1)
    pdf.ln(5)
    pdf.set_font('', '', 11)
    for j in range(20):
      y_before = pdf.get_y()
      pdf.multi_cell(19, 10, f'{j + 1}', border=1, align='C')
      pdf.set_xy(pdf.l_margin + 19, y_before)
      next_x = 19
      for i in range(10):
        value = REACTION_LABEL.loc[j, i] + '\n' + format(REACTION_TIME.loc[j, i], '.3f')
        pdf.multi_cell(15, 5, value, align='C', border=1)
        next_x += 15
        pdf.set_xy(next_x + pdf.l_margin, y_before)
      pdf.ln(10)

  DETAIL = [int(x) for x in df_data[:200]]
  TRUE_TIME = [float(x) for x in df_data[200:]]
  begin = [0, 16, 33, 50, 66, 83, 100, 116, 133, 150, 166, 183]
  end = [16, 33, 50, 66, 83, 100, 116, 133, 150, 166, 183, 200]
  CORRECT_REACTION = []
  INCORRECT_REACTION = []
  CORRECT_NON_REACTION = []
  INCORRECT_NON_REACTION = []
  def get_mean_time_correct_react(begin, end):
    TRUE_TIME_CORRECT = []
    TRUE_TIME_INCORRECT = []
    TRUE_TIME_CORRECT_NON_REACTION = 0
    TRUE_TIME_INCORRECT_NON_REACTION = 0
    for i in range(begin, end):
      if DETAIL[i] == 11:
        TRUE_TIME_CORRECT.append(TRUE_TIME[i])
      elif DETAIL[i] == 1:
        TRUE_TIME_INCORRECT.append(TRUE_TIME[i])
      elif DETAIL[i] == 0:
        TRUE_TIME_CORRECT_NON_REACTION += 1
      else: 
        TRUE_TIME_INCORRECT_NON_REACTION += 1
    CORRECT_REACTION.append(np.mean(TRUE_TIME_CORRECT))
    INCORRECT_REACTION.append(np.mean(TRUE_TIME_INCORRECT))
    CORRECT_NON_REACTION.append(TRUE_TIME_CORRECT_NON_REACTION + len(TRUE_TIME_CORRECT))
    INCORRECT_NON_REACTION.append(TRUE_TIME_INCORRECT_NON_REACTION + len(TRUE_TIME_INCORRECT))
  for i in range(12):
    get_mean_time_correct_react(begin[i], end[i])

  b, a = np.polyfit(np.arange(1, 13, 1), CORRECT_REACTION, deg=1)
  plt.rcParams["figure.figsize"] = (17, 13)
  plt.rcParams['axes.spines.right'] = False
  plt.rcParams['axes.spines.top'] = False
  fig, ax = plt.subplots(3, sharex=True)
  # First plot
  ax[0].plot(1, 0, ">k", transform=ax[0].get_yaxis_transform(), clip_on=False)
  ax[0].plot(-0.3, 1, "^k", transform=ax[0].get_xaxis_transform(), clip_on=False)
  ax[0].scatter(np.arange(0, 12, 1), CORRECT_REACTION, label = 'Thời gian "phản ứng chính xác" trung bình', color='b')
  ax[0].scatter(np.arange(0, 12, 1), INCORRECT_REACTION, label = 'Thời gian "phản ứng không chính xác" trung bình', facecolors='none', edgecolors='b')
  ax[0].plot(CORRECT_REACTION, color='b')
  ax[0].plot(INCORRECT_REACTION, color='b')
  ax[0].plot(a + b * np.arange(0, 12, 1), color='r')
  ax[0].set_xlim(-0.3, 11.5)
  ax[0].set_ylim(0, 1.6)
  ax[0].set_xticks(np.arange(0, 11.5, 1))
  ax[0].set_yticks(np.arange(0, 1.9, 0.2))
  ax[0].xaxis.set_ticklabels([])
  ax[0].set_yticklabels(['0.0', '0.2', '0.4', '0.6', '0.8', '1.0', '1.2', '1.4', '1.6', ''])
  ax[0].legend(loc=2, frameon=False)
  # Second plot
  ax[1].plot(1, 0, ">k", transform=ax[1].get_yaxis_transform(), clip_on=False)
  ax[1].plot(-0.3, 1, "^k", transform=ax[1].get_xaxis_transform(), clip_on=False)
  ax[1].bar(np.arange(0, 12, 1), CORRECT_NON_REACTION, width=0.1, color='b')
  ax[1].set_ylim(0, 23)
  ax[1].set_yticks(np.arange(0, 25, 5))
  ax[1].set_title('Tổng số "phản ứng chính xác" và "không phản ứng chính xác"', loc='left', fontsize=11)
  # Third plot
  ax[2].plot(1, 0, ">k", transform=ax[2].get_yaxis_transform(), clip_on=False)
  ax[2].plot(-0.3, 1, "^k", transform=ax[2].get_xaxis_transform(), clip_on=False)
  ax[2].bar(np.arange(0, 12, 1), INCORRECT_NON_REACTION, width=0.1, color='b')
  ax[2].set_ylim(0, 12)
  ax[2].set_yticks(np.arange(0, 15, 5))
  ax[2].set_title('Tổng số "phản ứng không chính xác" và "không phản ứng không chính xác"', loc='left', fontsize=11)
  ax[2].set_xlabel('Khoảng (1 khoảng = 30 giây)', loc='right', fontsize=11)
  ax[2].set_xticklabels([str(x) for x in range(1, 13)])
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
  pdf.cell(0, 5, 'Thử nghiệm nhận thức (Cognitrone - COG)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Thử nghiệm hiệu suất chung để đánh giá sự chú ý và tập trung')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Mẫu thử nghiệm 4 - Tương đương mẫu 1, thời gian thực hiện cố định cho mỗi mẫu')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, '20 mẫu với 10 kích thích mỗi mẫu (=200 kích thích/80 được yêu cầu)')
  pdf.ln(5)
  pdf.cell(0, 5, test_administration, border='B')
  pdf.ln(10)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Kết quả thử nghiệm - Mẫu chuẩn:')
  pdf.ln(5)
  output_df_to_pdf(pdf, FINAL_RESULT)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nhận xét: '), 4, 'Nhận xét: ', border='T')
  pdf.set_font('', '', 9)
  pdf.cell(0, 4, 'Tỷ lệ phân cấp (PR) và điểm T là kết quả của so sánh với toàn bộ mẫu so sánh tiêu chuẩn. Khoảng tin cậy được đưa ra', border='T')
  pdf.ln(4)
  pdf.cell(0, 4, 'trong ngoặc đơn bên cạnh các điểm so sánh có sai số là 5%.')
  pdf.ln(4)
  pdf.cell(0, 4, '(1) Không phản ứng không chính xác = nút không được nhấn ở kích thích cần thiết')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Nhận xét và giải thích về các biến thử nghiệm:')
  pdf.ln(8)
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Tổng số "phản ứng chính xác":', border='T')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Biến này cho biết tổng số mẫu có phản ứng chính xác. Điều này có nghĩa là nút được nhấn khi các số liệu giống nhau được trình bày và không có nút nào được nhấn khi các số liệu không giống nhau. Biến này là thước đo độ chính xác của quá trình kiểm tra dưới áp lực thời gian. Cụ thể hơn, nó phản ánh khả năng phân tích các mẫu một cách toàn diện trong thời hạn quy định.')
  pdf.multi_cell(0, 4, 'Do đó, xếp hạng phần trăm cao có nghĩa là người trả lời có khả năng tập trung sự chú ý của mình vào các vấn đề liên quan khi chịu áp lực về thời gian và do đó có thể làm việc một cách chính xác.',
                border='B')
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Tổng số "phản ứng không chính xác":')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Biến này cho biết tổng số mẫu có phản ứng không chính xác. Điều này có nghĩa là nút không được nhấn khi các số liệu giống hệt nhau hoặc nút được nhấn khi số liệu không giống nhau.')
  pdf.multi_cell(0, 4, 'Biến này phản ánh việc xác định không chính xác các điểm tương đồng do kết quả của việc kết thúc quá trình kiểm tra sớm mà nguyên nhân là do các điều kiện kiểm tra. Vì thời gian trình bày có hạn nên không thể phân tích đầy đủ các mục riêng lẻ; nhận thức rằng hai hình dạng giống nhau dẫn đến việc chúng được phân loại là giống hệt nhau.')
  pdf.multi_cell(0, 4, 'Do đó, xếp hạng phần trăm thấp cho thấy rằng dưới áp lực thời gian, người trả lời có xu hướng tập trung sự chú ý của mình vào những vấn đề không liên quan; do đó cô ấy/anh ấy phân tích chúng với độ chính xác thấp và hệ quả là đưa ra quyết định không chính xác.',
                border='B')
  pdf.set_font('', 'I', 9)
  pdf.multi_cell(0, 4, 'Lưu ý:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, '- Các nhận xét và giải thích được đề cập ở trên về các biến thử nghiệm có thể được bật hoặc tắt tại tab "Cài đặt Mở rộng" của Hệ thống Kiểm tra Vienna.')
  pdf.multi_cell(0, 4, '- Bạn có thể tìm thấy mô tả chi tiết của tất cả các biến thử nghiệm và toàn bộ các ghi chú giải thích trong sổ tay thử nghiệm kỹ thuật số (Sổ tay thử nghiệm này có thể được hiển thị và in ra thông qua giao diện người dùng của Hệ thống Thử nghiệm Vienna).')
  pdf.ln(8)
  pdf.add_page()
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Hồ sơ - Mẫu chuẩn:')
  pdf.ln(5)
  pdf.image(test_tmp_path, w=170)
  pdf.ln(8)
  pdf.cell(0, 5, 'Đồ thị tiến độ:')
  pdf.ln(5)
  pdf.image(test_protocol_temp_file, w=170)
  pdf.ln(5)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nhận xét: '), 4, 'Nhận xét: ')
  pdf.set_font('', '', 9)
  pdf.set_text_color(255, 0, 0)
  pdf.cell(pdf.get_string_width('─── '), 4, '─── ')
  pdf.set_text_color(0, 0, 0)
  pdf.cell(0, 4, 'Đường hồi quy cho thời gian trung bình "phản ứng đúng"')
  pdf.add_page()
  pdf.set_font('', 'B', 11)
  pdf.multi_cell(0, 8, 'Giao thức thử nghiệm:')
  output_df_to_pdf_protocol(pdf)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nhận xét: '), 4, 'Nhận xét: ')
  pdf.set_font('', '', 9)
  pdf.cell(0, 4, 'Các ký hiệu trong các trường có nghĩa là:? +: Không phản ứng đúng (không nhấn nút khi kích thích không cần thiết), ? -:')
  pdf.ln(4)
  pdf.multi_cell(0, 4, 'không chính xác không phản ứng (không nhấn nút ở một kích thích cần thiết) 1+: phản ứng đúng (nhấn nút ở một kích thích cần thiết), 1-: phản ứng không chính xác (nút được nhấn ở một kích thích không bắt buộc); Thời gian làm việc tính bằng giây; ——: Vật phẩm không được trình bày')
  pdf.output(output_path)
  os.remove(test_tmp_path)
  os.remove(test_protocol_temp_file)