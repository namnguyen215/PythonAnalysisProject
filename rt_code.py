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


  TEST_VARIABLE = ['Thời gian phản ứng trung bình (2)', 'Thời gian vận động trung bình (2)', 'Đo thời gian phản ứng phân tán (3)',
                  'Đo thời gian vận động phân tán (3)', 'Kết quả bổ sung', 'Phản ứng chính xác', 'Không có phản ứng',
                  'Phản ứng không hoàn chỉnh', 'Phản ứng không chính xác']
  RAW_SCORE = [int(info_df.loc[0, 'MRZ']), int(info_df.loc[0, 'MMZ']), int(info_df.loc[0, 'SDRZ']), int(info_df.loc[0, 'SDMZ']), '  ', 
              int(info_df.loc[0, 'RR']), int(info_df.loc[0, 'NR']), int(info_df.loc[0, 'UR']), int(info_df.loc[0, 'FR'])]
  PR = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).perc_rank() for i in range(4)]
  PR.extend(['  '] + [''] * 4)
  T_SCORE = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).interval_rt() for i in range(4)]
  T_SCORE.extend(['  '] + [''] * 4)
  FINAL_RESULT = pd.DataFrame(
      {'Biến thử nghiệm': TEST_VARIABLE,
      'Số liệu gốc': [str(i) for i in RAW_SCORE],
      'PR': PR,
      'T': T_SCORE}
      )
  plt.rcParams.update(plt.rcParamsDefault)
  plt.rcParams["figure.figsize"] = (8.5, 2)
  TEST_VARIABLE = ['Thời gian phản ứng trung bình', 'Thời gian vận động trung bình',
                  'Đo thời gian phản ứng phân tán', 'Đo thời gian vận động phân tán']
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
  STIM_TYPE = ['Vàng + Tông màu', 'Đỏ', 'Vàng', 'Tông màu', 'Đỏ + Tông màu', 'Vàng', 'Vàng + Tông màu', 'Đỏ', 'Vàng', 'Tông màu', 'Đỏ + Tông màu', 'Vàng', 'Vàng + Tông màu', 'Đỏ + Tông màu', 'Vàng', 'Vàng + Tông màu','Đỏ', 'Vàng + Tông màu', 'Vàng + Tông màu', 'Đỏ + Tông màu', 'Vàng', 'Vàng + Tông màu', 'Đỏ', 'Vàng + Tông màu','Vàng + Tông màu', 'Đỏ', 'Vàng', 'Tông màu', 'Đỏ + Tông màu', 'Vàng', 'Vàng + Tông màu', 'Đỏ','Vàng', 'Tông màu', 'Đỏ + Tông màu', 'Vàng', 'Vàng + Tông màu', 'Đỏ + Tông màu', 'Vàng', 'Vàng + Tông màu','Đỏ', 'Vàng + Tông màu', 'Vàng + Tông màu', 'Đỏ + Tông màu', 'Vàng', 'Vàng + Tông màu', 'Đỏ', 'Vàng + Tông màu']
  REQUIRED = ['Yes' if int(x) == 1 else 'No' for x in df_data[48:96]]
  EVALUATION = []
  for x in df_data[96:144]:
    # try:
      if int(x) == 1:
        EVALUATION.append('Phản ứng không chính xác')
      elif int(x) == 2:
        EVALUATION.append('Không phản ứng')
      elif int(x) == 3:
          EVALUATION.append('Phản ứng không hoàn chỉnh')
      elif int(x) == 4:
        EVALUATION.append('Phản ứng chính xác')
      else: EVALUATION.append('——')

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
  ax[0].set_title('Thời gian phản ứng tính bằng mili giây', loc='left', fontsize=11)
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
  ax[1].set_title('Thời gian vận động tính bằng mili giây', loc='left', fontsize=11)
  ax[1].set_xlabel('Kích thích', loc='right', fontsize=11)
  test_protocol_temp_file = read_help_file("test_protocol.png")
  plt.savefig(test_protocol_temp_file, format="png", bbox_inches='tight', dpi=200)

  REACTION_TIME = ['——' if math.isnan(x) else int(x) for x in REACTION_TIME]
  MOTOR_TIME = ['——' if math.isnan(x) else int(x) for x in MOTOR_TIME]
  for i in range(len(MOTOR_TIME)):
    if EVALUATION[i] == 'Phản ứng không chính xác':
      REACTION_TIME[i] = f'({REACTION_TIME[i]})'
      MOTOR_TIME[i] = f'({MOTOR_TIME[i]})'
  protocol = pd.DataFrame(
      {'STT \n ': STIM_NR,
      'Loại kích thích \n ': STIM_TYPE,
      'Yêu cầu \n ': REQUIRED,
      'Đánh giá \n ': EVALUATION,
      'Thời gian phản ứng (mili giây)': REACTION_TIME,
      'Thời gian vận động (mili giây)': MOTOR_TIME}
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
          if value in ['Kết quả bổ sung', '  ']:
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
  pdf.cell(0, 5, 'Kiểm tra phản ứng (Reaction Test - RT)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Thử nghiệm để đánh giá thời gian phản ứng đối với các kích thích bằng âm thanh và hình ảnh.')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Mẫu thử S3 - Màu vàng / âm phản ứng lựa chọn')
  pdf.ln(5)
  pdf.set_font('', '', 11)
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
  pdf.cell(0, 4, '(1) Tất cả các mục thời gian tính bằng mili giây')
  pdf.ln(4)
  pdf.cell(0, 4, '(2) Thời gian phản ứng trung bình sau khi chuẩn hóa Box-Cox')
  pdf.ln(4)
  pdf.cell(0, 4, '(3) Độ lệch chuẩn của thời gian phản ứng sau khi chuẩn hóa Box-Cox ')
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Nhận xét và giải thích về các biến thử nghiệm:')
  pdf.ln(8)
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Thời gian phản ứng trung bình:', border='T')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Khi sử dụng nút nghỉ, thời gian phản ứng là khoảng thời gian từ khi bắt đầu tác động kích thích liên quan đến thời điểm ngón tay rời khỏi nút nghỉ.')
  pdf.multi_cell(0, 4, 'Điểm này là thời gian phản ứng của mỗi lần thử nghiệm. Điểm cao cho thấy rằng so với dân số tham chiếu, người được hỏi có khả năng phản ứng với các kích thích hoặc chùm kích thích có liên quan nhanh trên mức trung bình.',
                border='B')
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Thời gian vận động trung bình:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Thời gian vận động là thời gian từ khi ngón tay rời khỏi nút nghỉ đến khi ngón tay tiếp xúc với nút phản ứng khi một kích thích liên quan xuất hiện.')
  pdf.multi_cell(0, 4, 'Điểm này cung cấp thông tin về tốc độ di chuyển của người trả lời. Điểm số cao cho thấy rằng so với dân số tham chiếu, người được hỏi có khả năng thực hiện một cách nhanh chóng quá trình hành động đã lên kế hoạch trong các tình huống phản ứng trên mức trung bình.',
                border='B')
  pdf.set_font('', 'I', 9)
  pdf.multi_cell(0, 4, 'Lưu ý:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, '- Các nhận xét và giải thích được đề cập ở trên về các biến thử nghiệm có thể được bật hoặc tắt tại tab "Cài đặt Mở rộng" của Hệ thống Kiểm tra Vienna.')
  pdf.multi_cell(0, 4, '- Bạn có thể tìm thấy mô tả chi tiết của tất cả các biến thử nghiệm và toàn bộ các ghi chú giải thích trong sổ tay thử nghiệm kỹ thuật số (Sổ tay thử nghiệm này có thể được hiển thị và in ra thông qua giao diện người dùng của Hệ thống Thử nghiệm Vienna).')
  pdf.add_page()
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Hồ sơ - Mẫu chuẩn:')
  pdf.ln(5)
  pdf.image(test_tmp_path, w=170)
  pdf.add_page()
  pdf.cell(0, 5, 'Đồ thị tiến độ:')
  pdf.ln(5)
  pdf.image(test_protocol_temp_file, w=170)
  pdf.ln(5)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nhận xét: '), 4, 'Nhận xét: ')
  pdf.set_font('', '', 9)
  pdf.set_text_color(0, 0, 255)
  pdf.cell(pdf.get_string_width('● '), 4, '● ')
  pdf.set_text_color(0, 0, 0)
  pdf.cell(pdf.get_string_width('Thời gian phản ứng với mỗi kích thích;  '), 4, 'Thời gian phản ứng với mỗi kích thích; ')
  pdf.set_text_color(255, 0, 0)
  pdf.cell(pdf.get_string_width('—— '), 4, '—— ')
  pdf.set_text_color(0, 0, 0)
  pdf.cell(0, 4, 'Đường hồi quy')
  pdf.add_page()
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Giao thức thử nghiệm:')
  pdf.ln(5)
  output_df_to_pdf_protocol(pdf, protocol)
  pdf.output(output_path)
  os.remove(test_tmp_path)
  os.remove(test_protocol_temp_file)

# set_path('C:/Users/NguyenNam/Desktop/Pilot_Analysis/bach cong thanh/RT03.ASC','result.pdf')