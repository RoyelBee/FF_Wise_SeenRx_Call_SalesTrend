import time
import pandas._libs.tslibs.base
start_time = time.time()

# # ------------- Generate all Data set ------------------------
import AdditonalFiles.generate_raw_data as data
print('Raw Data Generating...')
# data.sales_achiv_trend_data()
# data.seen_rx_data()
# data.doctor_call_data()

print('All Raw data created \n')
# # --------------------- Send All Data -----------------------
import AdditonalFiles.send_all_mail as all_mail
all_mail.send_all_report()

# ----------- Send Single RSM Mail ----------------------------
# import AdditonalFiles.send_rsm_mail as mail

# mail.send_report('CBU', 'rejaul.islam@transcombd.com')  # 'abul.basher@skf.transcombd.com'
# mail.send_report('CCF','bijoyskf1980@gmail.com')  # 'bijoyskf1980@gmail.com'
# mail.send_report('CCX', 'shafi.uddin@skf.transcombd.com')  # 'shafi.uddin@skf.transcombd.com'
# mail.send_report('CNH', 'abu.sayeed@skf.transcombd.com')  # 'abu.sayeed@skf.transcombd.com'
# mail.send_report('CKJ','ahmed.manjur@skf.transcombd.com')  # 'ahmed.manjur@skf.transcombd.com'
# mail.send_report('CMT','mizanur.khan@skf.transcombd.com')  # 'mizanur.khan@skf.transcombd.com'
# mail.send_report('CRB','almamun.abdullah@skf.transcombd.com')  # 'almamun.abdullah@skf.transcombd.com'
# mail.send_report('CRP','sanaullah.khan@skf.transcombd.com')  # 'sanaullah.khan@skf.transcombd.com' prb
# mail.send_report('CSB','masud.rana@skf.transcombd.com')  # 'masud.rana@skf.transcombd.com'
# mail.send_report('CMT','mizanur.khan@skf.transcombd.com')  # 'mizanur.khan@skf.transcombd.com'

print('Time takes = ', round((time.time() - start_time) / 60, 2), 'Min')
