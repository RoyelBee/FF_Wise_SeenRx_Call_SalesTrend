import time

# start_time = time.time()
# import AdditonalFiles.send_mail as mail
#
# mail.send_report()
#
#
# end_time = time.time()
# print('Time Takes = ', end_time - start_time)

# # Generate all Data set ------------------------
import AdditonalFiles.generate_raw_data as data
# data.seen_rx_data()
# data.doctor_call_data()
# data.sales_trend_data()

import AdditonalFiles.send_mail as mail

mail.send_report("CBU")
