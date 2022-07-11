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

  sample_data = pd.read_excel(sample_path, sheet_name='SMK')
  sample_data = sample_data[['MEANA', 'MEANH', 'MEANV', 'DEVA', 'DEVH', 'DEVV', 
                            'IDEAL', 'MEANA5', 'MEANH5', 'MEANV5', 'DEVA5',
                            'DEVH5', 'DEVV5', 'IDEAL5']]
  user_data = pd.read_csv(user_path, sep='_', header=None)

  df_info = user_data.iloc[0:6, :-1].stack().tolist()
  col = ['PCODE', 'SEX', 'AGE', 'AGEMON', 'EDLEV', 'ADDINF1', 'KBZ', 'FORM', 'VERSION', 'SERIAL', 'SOURCE', 'AWCODE',
        'DATE', 'TIME1', 'TIME2', 'IDNO', '?', 'DEVICE', 'CANCEL', '?', 'MEANA', 'MEANH', 'MEANV', 'DEVA', 'DEVH', 
        'DEVV', 'IDEAL', 'SUMA', 'SUMH', 'SUMV', 'MEANA5', 'MEANH5', 'MEANV5', 'DEVA5', 'DEVH5',
        'DEVV5', 'IDEAL5', 'SUMA5', 'SUMH5', 'SUMV5']
  df_data = user_data.iloc[6:,:-4].stack().tolist()
  info_df = pd.DataFrame([df_info], columns=col)
  info_df.replace(r'^\s*$', np.nan, regex=True, inplace=True)



  TEST_VARIABLE = ['Góc lệch trung bình (tính bằng độ)', 'Độ lệch ngang trung bình (tính bằng pixels)', 'Độ lệch dọc trung bình (tính bằng pixels)',
                  'Độ lệch góc phân tán (tính bằng độ)', 'Độ lệch ngang phân tán (tính bằng pixel)', 'Độ lệch dọc phân tán (tính bằng pixel)',
                  'Thời gian trong khung lí tưởng (tính bằng %)', 'Kết quả tạm thời sau 5 phút', 'Góc lệch trung bình (tính bằng độ)', 
                  'Độ lệch ngang trung bình (tính bằng pixels)', 'Độ lệch dọc trung bình (tính bằng pixels)', 'Độ lệch góc phân tán (tính bằng độ)', 
                  'Độ lệch ngang phân tán (tính bằng pixel)', 'Độ lệch dọc phân tán (tính bằng pixel)', 'Thời gian trong khung lí tưởng (tính bằng %)']
  RAW_SCORE = [float(info_df.loc[0, 'MEANA']), float(info_df.loc[0, 'MEANH']), float(info_df.loc[0, 'MEANV']),
              float(info_df.loc[0, 'DEVA']), float(info_df.loc[0, 'DEVH']), float(info_df.loc[0, 'DEVV']), int(info_df.loc[0, 'IDEAL']), 
              float(info_df.loc[0, 'MEANA5']), float(info_df.loc[0, 'MEANH5']), float(info_df.loc[0, 'MEANV5']),
              float(info_df.loc[0, 'DEVA5']), float(info_df.loc[0, 'DEVH5']), float(info_df.loc[0, 'DEVV5']), int(info_df.loc[0, 'IDEAL5'])]
  PR = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).perc_rank() for i in range(14)]
  PR.insert(7, '  ')
  T_SCORE = [compute_score(sample_data.iloc[:, i], RAW_SCORE[i]).interval_rt() for i in range(14)]
  T_SCORE.insert(7, '  ')
  RAW_SCORE.insert(7, '  ')
  FINAL_RESULT = pd.DataFrame(
      {'Biến thử nghiệm': TEST_VARIABLE,
      'Số liệu gốc': RAW_SCORE,
      'PR': PR,
      'T': T_SCORE}
      )

  PARTIAL_INTERVAL = list(range(1, 12))
  WA_MEAN = [round(float(x), 1) for x in df_data[:11]]
  HA_MEAN = [round(float(x), 1) for x in df_data[11: 22]]
  VA_MEAN = [round(float(x), 1) for x in df_data[22: 33]]
  WA_DISP = [round(float(x), 1) for x in df_data[33: 44]]
  HA_DISP = [round(float(x), 1) for x in df_data[44: 55]]
  VA_DISP = [round(float(x), 1) for x in df_data[55: 66]]
  ID = [int(x) for x in df_data[66:]]
  protocol = pd.DataFrame(
      {'Khoảng \n một phần \n ': PARTIAL_INTERVAL,
      'WA \n (Trung bình)': WA_MEAN,
      'HA \n (Trung bình)': HA_MEAN,
      'VA \n (Trung bình)': VA_MEAN,
      'WA \n (Phân tán) \n ': WA_DISP,
      'HA \n (Phân tán) \n ': HA_DISP,
      'VA \n (Phân tán) \n ': VA_DISP,
      'Thời gian trong khung lý tưởng': ID}
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
          if value == 'Kết quả tạm thời sau 5 phút':
            pdf.set_font('', 'B', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='T')
          elif value == '  ':
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='T')
          else:
            pdf.set_font('', '', 11)
            pdf.cell(table_cell_width[i], table_cell_height, value, align=align, border='')
          i += 1
        pdf.ln(table_cell_height)

  def output_df_to_pdf_protocol(pdf, df):
      # A cell is a rectangular area, possibly framed, which contains some text
      # Set the width and height of cell
      table_cell_width = [0, 22, 19, 20, 20, 21, 21, 21, 28]
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
      pdf.ln(table_cell_height * 3)
      # Select a font as Times, regular, 11
      pdf.set_font('times', '', 11)
      # Loop over to print each data in the table
      for row in df.itertuples():
        i = 1
        for col in ['_1', '_2', '_3', '_4', '_5', '_6', '_7', '_8']:
          value = str(getattr(row, col))
          pdf.cell(table_cell_width[i], table_cell_height, value, align='C', border=1)
          i += 1
        pdf.ln(table_cell_height)

  plt.rcdefaults()
  x_plot = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() for i in range(7)]))
  y_plot = list(reversed(TEST_VARIABLE[:7]))
  x_error = list(reversed([compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='main').interval_rt() - compute_score(sample_data.iloc[:, i], RAW_SCORE[i], ci='lower').interval_rt() for i in range(7)]))
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
  pdf.cell(0, 5, 'Kiểm tra phối hợp cảm biến vận động (Sensomotor Coordination - SMK)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, 'Thử nghiệm để kiểm tra phối hợp cảm biến vận động')
  pdf.ln(5)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Mẫu thử nghiệm S1 - Mẫu ngắn (thời gian: 10 phút)')
  pdf.ln(5)
  pdf.set_font('', '', 11)
  pdf.cell(0, 5, test_administration, border='B')
  pdf.ln(10)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Kết quả thử nghiệm - Người trưởng thành:')
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
  pdf.multi_cell(0, 4, 'Thời gian trong khung lí tưởng (tính bằng %):', border='T')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, 'Trong một khoảng thời gian nhất định, phần trăm thời gian mà đoạn tròn nằm trong khung lý tưởng (100% = đoạn tròn luôn nằm trong khung lý tưởng). Phạm vi lý tưởng được định nghĩa là độ lệch tối đa +/- 25 pixels trong chuyển động ngang và dọc và độ lệch tối đa +/- 25 độ trong chuyển động nghiêng.')
  pdf.multi_cell(0, 4, 'Trên thực tế, hiếm khi người trả lời đạt điểm trung bình hoặc tốt cho tất cả các biến chính hoặc phụ mà bị thâm hụt một biến hoặc ngược lại. Nếu tình huống này phát sinh, quy trình kiểm tra trước tiên cần được phân tích để xác định điểm mà người trả lời cho thấy điểm yếu trong thử nghiệm. Nếu vấn đề nảy sinh ngay từ đầu, có thể người trả lời đã không hiểu các hướng dẫn.')
  pdf.multi_cell(0, 4, 'Trong những trường hợp khác, sự mệt mỏi hoặc thiếu tập trung có thể là một phần nguyên nhân. Trong cả hai trường hợp, thử nghiệm phải được lặp lại (nếu có thể) dưới sự giám sát của người quản lý thử nghiệm. Do đó, việc phân tích các kết quả trung gian có thể cho biết thời điểm hiệu suất của cá nhân đạt đến đỉnh điểm hoặc tăng lên, một cách độc lập với điểm chuẩn.',
                border='B')
  pdf.set_font('', 'I', 9)
  pdf.multi_cell(0, 4, 'Lưu ý:')
  pdf.set_font('', '', 9)
  pdf.multi_cell(0, 4, '- Các nhận xét và giải thích được đề cập ở trên về các biến thử nghiệm có thể được bật hoặc tắt tại tab "Cài đặt Mở rộng" của Hệ thống Kiểm tra Vienna.')
  pdf.multi_cell(0, 4, '- Bạn có thể tìm thấy mô tả chi tiết của tất cả các biến thử nghiệm và toàn bộ các ghi chú giải thích trong sổ tay thử nghiệm kỹ thuật số (Sổ tay thử nghiệm này có thể được hiển thị và in ra thông qua giao diện người dùng của Hệ thống Thử nghiệm Vienna).')
  pdf.add_page()
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Hồ sơ - người trưởng thành:')
  pdf.ln(5)
  pdf.image(test_tmp_path, w=170)
  pdf.ln(10)
  pdf.set_font('', 'B', 11)
  pdf.cell(0, 5, 'Giao thức thử nghiệm:')
  pdf.ln(8)
  output_df_to_pdf_protocol(pdf, protocol)
  pdf.set_font('', 'I', 9)
  pdf.cell(pdf.get_string_width('Nhận xét: '), 4, 'Nhận xét: ', border='T')
  pdf.set_font('', '', 9)
  pdf.cell(0, 4, 'Khoảng một phần = một phần tương ứng với 56 giây (800 khung hình); WA = độ lệch góc theo độ; HA = Độ lệch ngang tính')
  pdf.ln(4)
  pdf.multi_cell(0, 4, 'bằng pixel; VA = Độ lệch dọc tính bằng pixel (0 = không lệch); Thời gian trong khung lý tưởng = phần trăm, khoảng thời gian mà đoạn tròn nằm trong khung lý tưởng trong một khoảng thời gian riêng phần (100% = đoạn tròn luôn nằm trong khung lý tưởng)')
  pdf.output(output_path)
  os.remove(test_tmp_path)