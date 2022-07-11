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



  TEST_VARIABLE = ['Chế độ thích ứng kết quả tổng thể (thời lượng thử nghiệm: 4 phút)', 'Chính xác', 'Không chính xác',
                  'Bị bỏ qua', 'Thời gian phản ứng trung vị', 'Số lượng kích thích', 'Các phản ứng']
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
      {'Biến thử nghiệm': TEST_VARIABLE,
      'Số liệu gốc': [str(i) for i in RAW_SCORE],
      'PR': PR,
      'T': T_SCORE}
      )

  plt.rcParams.update(plt.rcParamsDefault)
  plt.rcParams["figure.figsize"] = (8.5, 2)
  TEST_VARIABLE = ['Chính xác', 'Không chính xác', 'Bị bỏ qua']
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
      # Select a font as times, regular, 11
      pdf.set_font('times', '', 11)
      # Loop over to print each data in the table
      for row in df.itertuples():
        i = 0
        for col in ['_1', '_2', 'PR', 'T']:
          align = '' if col == '_1' else 'C'
          value = str(getattr(row, col))
          if value == 'Chế độ thích ứng kết quả tổng thể (thời lượng thử nghiệm: 4 phút)':
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
  ax[0].set_title('Thời gian phản ứng (tính bằng mili giây)', loc='left', fontsize=11)
  # Second plot
  ax[1].plot(1, 0, ">k", transform=ax[1].get_yaxis_transform(), clip_on=False)
  ax[1].plot(-5, 1, "^k", transform=ax[1].get_xaxis_transform(), clip_on=False)
  ax[1].bar(np.arange(1, NBR_STIMULI+1, 1), OMITTED, width=2, color='b')
  ax[1].set_ylim(0, 1.3)
  ax[1].set_yticks(np.arange(0, 1.3, 1))
  ax[1].set_title('Bị bỏ qua', loc='left', fontsize=11)
  # Third plot
  ax[2].plot(1, 0, ">k", transform=ax[2].get_yaxis_transform(), clip_on=False)
  ax[2].plot(-5, 1, "^k", transform=ax[2].get_xaxis_transform(), clip_on=False)
  ax[2].bar(np.arange(1, NBR_STIMULI+1, 1), INCORRECT, width=2, color='b')
  ax[2].set_ylim(0, 1.3)
  ax[2].set_yticks(np.arange(0, 1.3, 1))
  ax[2].set_title('Không chính xác', loc='left', fontsize=11)
  ax[2].set_xlabel('Kích thích', loc='right', fontsize=11)
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
  pdf.cell(0, 5, 'Thử nghiệm xác định (Determination Test - DT)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Thí nghiệm phản ứng trắc nghiệm có kích thích phức tạp')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Mẫu thử nghiệm S1 - Mẫu ngắn có trình bày kích thích thích ứng (tất cả các loại kích thích)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Thử nghiệm kéo dài 4 phút')
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
  pdf.ln(8)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Nhận xét và giải thích về các biến thử nghiệm:')
  pdf.ln(8)
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Chính xác:', border='T')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Biến chính này nêu chi tiết tổng số phản ứng chính xác được thực hiện trước khi bắt đầu hành động tiếp theo trừ một kích thích. Nó đo lường khả năng tiếp tục phản ứng nhanh chóng và thích hợp trong chuỗi phản ứng của người trả lời để, bao gồm cả khi làm việc gần đến giới hạn chịu đựng căng thẳng của cá nhân họ.',
                border='B')
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Không chính xác:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Biến này mô tả xu hướng nhầm lẫn giữa các phản ứng khác nhau. Phản ứng không chính xác phát sinh bởi vì khi bị căng thẳng, người trả lời không thể bảo vệ phản ứng thích hợp khỏi ảnh hưởng của các kích thích không thích hợp.',
                border='B')
  pdf.set_font('', 'B', 9)
  pdf.multi_cell(0, 4, 'Bị bỏ qua:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Biến phụ này cho biết liệu các phản hồi có bị bỏ qua dưới áp lực thời gian hay không. Những cá nhân có điểm rất cao về biến này có xu hướng không thể duy trì sự chú ý của họ khi thực hiện các nhiệm vụ thuộc loại này trong tình trạng căng thẳng; điều này có nghĩa là trong những tình huống căng thẳng, họ có thể có xu hướng bỏ cuộc.',
                border='B')
  pdf.set_font('', 'I', 9)
  pdf.multi_cell(0, 4, 'Lưu ý:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, '- Các nhận xét và giải thích được đề cập ở trên về các biến thử nghiệm có thể được bật hoặc tắt tại tab "Cài đặt Mở rộng" của Hệ thống Kiểm tra Vienna.')
  pdf.multi_cell(0, 4, '- Bạn có thể tìm thấy mô tả chi tiết của tất cả các biến thử nghiệm và toàn bộ các ghi chú giải thích trong sổ tay thử nghiệm kỹ thuật số (Sổ tay thử nghiệm này có thể được hiển thị và in ra thông qua giao diện người dùng của Hệ thống Thử nghiệm Vienna).')
  pdf.ln(8)
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
  pdf.set_text_color(255, 0, 0)
  pdf.cell(pdf.get_string_width('—— '), 4, '—— ')
  pdf.set_text_color(0, 0, 0)
  pdf.cell(0, 4, 'Đường hồi quy')
  pdf.output(output_path)
  os.remove(test_tmp_path)
  os.remove(test_protocol_temp_file)
  