import pandas as pd
from PIL import ImageDraw, ImageFont
from PIL import Image
import AdditonalFiles.db_connection as conn
from datetime import datetime
from datetime import date, timedelta
import warnings
from statistics import mean
from datetime import datetime
from datetime import date
from calendar import monthrange

warnings.filterwarnings('ignore')

today = datetime.today()
yesterday = date.today() - timedelta(days=1)

day = str(yesterday.day) + '-' + str(yesterday.strftime("%b")) + '-' + str(yesterday.year)

curr_month = str(today.strftime("%b")) + '-' + str(today.year)

font = ImageFont.truetype("./font/ROCK.ttf", 22, encoding="unic")
font1 = ImageFont.truetype("./font/ROCK.ttf", 20, encoding="unic")


def number_decorator(number):
    number = round(number, 2)
    number = format(number, ',')
    return number


def create_total_kpi():
    font = ImageFont.truetype("./font/ROCK.ttf", 40, encoding="unic")
    f = ImageFont.truetype("./font/ROCK.ttf", 25, encoding="unic")

    img = Image.open("./Images/total_kpi.png")
    img_draw = ImageDraw.Draw(img)

    trgdf = pd.read_sql_query(""" select sum(amount) as brand_target from V_FF_Brand_Target
            where brand in ( 'LIGAZID', 'EMAZID', 'LIPICON', 'AGLIP', 'CIFIBET',
                             'AMLEVO', 'CARDOBIS', 'RIVAROX', 'NOCLOG', 'BEMPID',
                             'AROTIDE', 'FOBUNID') """, conn.azure)

    target = int(trgdf.brand_target)

    salesdf = pd.read_sql_query(""" DECLARE @LastDay as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)-1
                DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), GETDATE()-1, 112)
                select sum(extinvmisc) as brandsales from V_FF_Brand_Sales where left(transdate,6)=@This_month and  TRANSDATE<=@LastDay and 
                brand in ( 'LIGAZID', 'EMAZID', 'LIPICON', 'AGLIP', 'CIFIBET',
                                 'AMLEVO', 'CARDOBIS', 'RIVAROX', 'NOCLOG', 'BEMPID',
                                 'AROTIDE', 'FOBUNID')  """, conn.azure)
    sales = int(salesdf.brandsales)

    if target > 1:
        sales_achivPercent = round((sales / target) * 100, 2)
    else:
        sales_achivPercent = 0
    #
    today = str(date.today())
    datee = datetime.strptime(today, "%Y-%m-%d")
    date_val = int(datee.day)
    num_days = monthrange(datee.year, datee.month)[1]

    if target > 1:
        trend_val = (sales / date_val) * num_days
        trend_percent = round((trend_val / target) * 100, 2)
    else:
        trend_percent = 0

    # # ------------------------- Seen Rx -----------------------------------
    seen_df = pd.read_csv('./Data/SeenRx/rsm_seen_total.csv')
    LIGAZID_seen = sum(seen_df.LIGAZID.tolist())
    EMAZID_seen = sum(seen_df.EMAZID.tolist())
    LIPICON_seen = sum(seen_df.LIPICON.tolist())
    AGLIP_seen = sum(seen_df.AGLIP.tolist())
    CIFIBET_seen = sum(seen_df.CIFIBET.tolist())
    AMLEVO_seen = sum(seen_df.AMLEVO.tolist())
    CARDOBIS_seen = sum(seen_df.CARDOBIS.tolist())
    RIVAROX_seen = sum(seen_df.RIVAROX.tolist())
    NOCLOG_seen = sum(seen_df.NOCLOG.tolist())
    BEMPID_seen = sum(seen_df.BEMPID.tolist())
    AROTIDE_seen = sum(seen_df.AROTIDE.tolist())
    FOBUNID_seen = sum(seen_df.FOBUNID.tolist())
    all_seenrx = int(sum([LIGAZID_seen, EMAZID_seen, LIPICON_seen, AGLIP_seen, CIFIBET_seen,
                          AMLEVO_seen, CARDOBIS_seen, RIVAROX_seen, NOCLOG_seen, BEMPID_seen, AROTIDE_seen,
                          FOBUNID_seen]))

    # # ---------------------- Doctor Call -----------------------------
    call_df = pd.read_csv('./Data/Call/rsm_call_total.csv')
    LIGAZID_call = sum(call_df.LIGAZID.tolist())
    EMAZID_call = sum(call_df.EMAZID.tolist())
    LIPICON_call = sum(call_df.LIPICON.tolist())
    AGLIP_call = sum(call_df.AGLIP.tolist())
    CIFIBET_call = sum(call_df.CIFIBET.tolist())
    AMLEVO_call = sum(call_df.AMLEVO.tolist())
    CARDOBIS_call = sum(call_df.CARDOBIS.tolist())
    RIVAROX_call = sum(call_df.RIVAROX.tolist())
    NOCLOG_call = sum(call_df.NOCLOG.tolist())
    BEMPID_call = sum(call_df.BEMPID.tolist())
    AROTIDE_call = sum(call_df.AROTIDE.tolist())
    FOBUNID_call = sum(call_df.FOBUNID.tolist())
    all_call = int(sum([LIGAZID_call, EMAZID_call, LIPICON_call, AGLIP_call, CIFIBET_call,
                        AMLEVO_call, CARDOBIS_call, RIVAROX_call, NOCLOG_call, BEMPID_call, AROTIDE_call,
                        FOBUNID_call]))

    img_draw.text((130, 22), '(' + str(day) + ')', (255, 255, 255), f)
    img_draw.text((90, 120), number_decorator(sales_achivPercent) + '%', (0, 0, 0), font)
    img_draw.text((410, 120), number_decorator(trend_percent) + '%', (0, 0, 0), font)
    img_draw.text((730, 120), number_decorator(all_seenrx), (0, 0, 0), font)
    img_draw.text((1020, 120), number_decorator(all_call), (0, 0, 0), font)

    img.save("./Images/total_kpi_images.png")
    print('2. Specefic KPI Created')


def create_rsm_total_kpi(rsm):
    font = ImageFont.truetype("./font/ROCK.ttf", 40, encoding="unic")
    f = ImageFont.truetype("./font/ROCK.ttf", 25, encoding="unic")

    img = Image.open("./Images/total_kpi.png")
    img_draw = ImageDraw.Draw(img)

    trgdf = pd.read_sql_query(""" select sum(amount) as brand_target from V_FF_Brand_Target
            where rsmtr like ? and brand in ( 'LIGAZID', 'EMAZID', 'LIPICON', 'AGLIP', 'CIFIBET',
                            'AMLEVO', 'CARDOBIS', 'RIVAROX', 'NOCLOG', 'BEMPID',
                            'AROTIDE', 'FOBUNID') """, conn.azure, params={rsm})

    target = int(trgdf.brand_target)

    salesdf = pd.read_sql_query(""" DECLARE @LastDay as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)-1
            DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), GETDATE()-1, 112)
            select sum(extinvmisc) as brandsales from V_FF_Brand_Sales where left(MSOTR, 3) like ? and left(transdate,
            6)=@This_month and  TRANSDATE<=@LastDay and
            brand in ( 'LIGAZID', 'EMAZID', 'LIPICON', 'AGLIP', 'CIFIBET',
            'AMLEVO', 'CARDOBIS', 'RIVAROX', 'NOCLOG', 'BEMPID',
            'AROTIDE', 'FOBUNID')   """, conn.azure, params={rsm})
    sales = int(salesdf.brandsales)

    if target > 1:
        sales_achivPercent = round((sales / target) * 100, 2)
    else:
        sales_achivPercent = 0
    #
    today = str(date.today())
    datee = datetime.strptime(today, "%Y-%m-%d")
    date_val = int(datee.day)
    num_days = monthrange(datee.year, datee.month)[1]

    if target > 1:
        trend_val = (sales / date_val) * num_days
        trend_percent = round((trend_val / target) * 100, 2)
    else:
        trend_percent = 0

    # # # ------------------------- Seen Rx -----------------------------------
    seen_df = pd.read_csv('./Data/SeenRx/rsm_seen_total.csv')
    seen_df.where(seen_df['FFTR'] == rsm, inplace=True)
    total_seenrx = int(seen_df[seen_df['FFTR'].notnull()].sum(axis=1))
    

    # # # ---------------------- Doctor Call -----------------------------
    call_df = pd.read_csv('./Data/Call/rsm_call_total.csv')
    call_df.where(call_df['FFTR'] == rsm, inplace=True)
    total_call = int(call_df[call_df['FFTR'].notnull()].sum(axis=1))

    #
    img_draw.text((130, 22), '(' + str(day) + ')', (255, 255, 255), f)
    img_draw.text((90, 120), number_decorator(sales_achivPercent) + '%', (0, 0, 0), font)
    img_draw.text((410, 120), number_decorator(trend_percent) + '%', (0, 0, 0), font)
    img_draw.text((750, 120), number_decorator(total_seenrx), (0, 0, 0), font)
    img_draw.text((1050, 120), number_decorator(total_call), (0, 0, 0), font)

    img.save("./Images/"+str(rsm)+"_total_kpi_images.png")
    print('2. Total Summary KPI Generated')


def all_kpi_images():
    img1 = Image.open("./Images/new_file.png")
    img_draw = ImageDraw.Draw(img1)

    sales_achiv = pd.read_excel('./Data/SalesAchivment/sales_achiv.xlsx')
    LIGZID_achiv = sales_achiv.LIGAZID.tolist()[0]
    EMAZID_achiv = sales_achiv.EMAZID.tolist()[0]
    LIPICON_achiv = sales_achiv.LIPICON.tolist()[0]
    AGLIP_achiv = sales_achiv.AGLIP.tolist()[0]
    CIFIBET_achiv = sales_achiv.CIFIBET.tolist()[0]
    AMLEVO_achiv = sales_achiv.AMLEVO.tolist()[0]
    CARDOBIS_achiv = sales_achiv.CARDOBIS.tolist()[0]
    RIVAROX_achiv = sales_achiv.RIVAROX.tolist()[0]
    NOCLOG_achiv = sales_achiv.NOCLOG.tolist()[0]
    BEMPID_achiv = sales_achiv.BEMPID.tolist()[0]
    AROTIDE_achiv = sales_achiv.AROTIDE.tolist()[0]
    FOBUNID_achiv = sales_achiv.FOBUNID.tolist()[0]

    img_draw.text((265, 13), 'up to (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((25, 85), number_decorator(LIGZID_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((120, 85), number_decorator(EMAZID_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((217, 85), number_decorator(LIPICON_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((320, 85), number_decorator(AGLIP_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((417, 85), number_decorator(CIFIBET_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((517, 85), number_decorator(AMLEVO_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((617, 85), number_decorator(CARDOBIS_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((717, 85), number_decorator(RIVAROX_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((817, 85), number_decorator(NOCLOG_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((917, 85), number_decorator(BEMPID_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((1017, 85), number_decorator(AROTIDE_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((1117, 85), number_decorator(FOBUNID_achiv) + '%', (0, 0, 0), font1)

    # # ------------------------------------------------------------------------------
    sales_trend = pd.read_excel('./Data/SalesTrend/sales_trend_data.xlsx')
    LIGZID_trend = sales_trend.LIGAZID.tolist()[0]
    EMAZID_trend = sales_trend.EMAZID.tolist()[0]
    LIPICON_trend = sales_trend.LIPICON.tolist()[0]
    AGLIP_trend = sales_trend.AGLIP.tolist()[0]
    CIFIBET_trend = sales_trend.CIFIBET.tolist()[0]
    AMLEVO_trend = sales_trend.AMLEVO.tolist()[0]
    CARDOBIS_trend = sales_trend.CARDOBIS.tolist()[0]
    RIVAROX_trend = sales_trend.RIVAROX.tolist()[0]
    NOCLOG_trend = sales_trend.NOCLOG.tolist()[0]
    BEMPID_trend = sales_trend.BEMPID.tolist()[0]
    AROTIDE_trend = sales_trend.AROTIDE.tolist()[0]
    FOBUNID_trend = sales_trend.FOBUNID.tolist()[0]

    img_draw.text((180, 130), 'up to (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((25, 200), number_decorator(LIGZID_trend) + '%', (0, 0, 0), font1)
    img_draw.text((120, 200), number_decorator(EMAZID_trend) + '%', (0, 0, 0), font1)
    img_draw.text((217, 200), number_decorator(LIPICON_trend) + '%', (0, 0, 0), font1)
    img_draw.text((320, 200), number_decorator(AGLIP_trend) + '%', (0, 0, 0), font1)
    img_draw.text((417, 200), number_decorator(CIFIBET_trend) + '%', (0, 0, 0), font1)
    img_draw.text((517, 200), number_decorator(AMLEVO_trend) + '%', (0, 0, 0), font1)
    img_draw.text((617, 200), number_decorator(CARDOBIS_trend) + '%', (0, 0, 0), font1)
    img_draw.text((717, 200), number_decorator(RIVAROX_trend) + '%', (0, 0, 0), font1)
    img_draw.text((817, 200), number_decorator(NOCLOG_trend) + '%', (0, 0, 0), font1)
    img_draw.text((917, 200), number_decorator(BEMPID_trend) + '%', (0, 0, 0), font1)
    img_draw.text((1017, 200), number_decorator(AROTIDE_trend) + '%', (0, 0, 0), font1)
    img_draw.text((1117, 200), number_decorator(FOBUNID_trend) + '%', (0, 0, 0), font1)

    # # ------------------------------ Seen Rx ----------------------------------------
    seen_rx = pd.read_excel('./Data/SeenRx/Seen_Rx_Data.xlsx')
    LIGZID_rx = seen_rx.LIGAZID.tolist()[0]
    EMAZID_rx = seen_rx.EMAZID.tolist()[0]
    LIPICON_rx = seen_rx.LIPICON.tolist()[0]
    AGLIP_rx = seen_rx.AGLIP.tolist()[0]
    CIFIBET_rx = seen_rx.CIFIBET.tolist()[0]
    AMLEVO_rx = seen_rx.AMLEVO.tolist()[0]
    CARDOBIS_rx = seen_rx.CARDOBIS.tolist()[0]
    RIVAROX_rx = seen_rx.RIVAROX.tolist()[0]
    NOCLOG_rx = seen_rx.NOCLOG.tolist()[0]
    BEMPID_rx = seen_rx.BEMPID.tolist()[0]
    AROTIDE_rx = seen_rx.AROTIDE.tolist()[0]
    FOBUNID_rx = seen_rx.FOBUNID.tolist()[0]

    img_draw.text((120, 247), '(' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((40, 325), number_decorator(LIGZID_rx), (0, 0, 0), font1)
    img_draw.text((135, 325), number_decorator(EMAZID_rx), (0, 0, 0), font1)
    img_draw.text((235, 325), number_decorator(LIPICON_rx), (0, 0, 0), font1)
    img_draw.text((335, 325), number_decorator(AGLIP_rx), (0, 0, 0), font1)
    img_draw.text((435, 325), number_decorator(CIFIBET_rx), (0, 0, 0), font1)
    img_draw.text((535, 325), number_decorator(AMLEVO_rx), (0, 0, 0), font1)
    img_draw.text((635, 325), number_decorator(CARDOBIS_rx), (0, 0, 0), font1)
    img_draw.text((735, 325), number_decorator(RIVAROX_rx), (0, 0, 0), font1)
    img_draw.text((835, 325), number_decorator(NOCLOG_rx), (0, 0, 0), font1)
    img_draw.text((935, 325), number_decorator(BEMPID_rx), (0, 0, 0), font1)
    img_draw.text((1035, 325), number_decorator(AROTIDE_rx), (0, 0, 0), font1)
    img_draw.text((1135, 325), number_decorator(FOBUNID_rx), (0, 0, 0), font1)

    # # ------------------------------ Doctor Call -----------------------------------
    doctor_call = pd.read_excel('./Data/Call/doctor_call_data.xlsx')
    LIGZID_call = doctor_call.LIGAZID.tolist()[0]
    EMAZID_call = doctor_call.EMAZID.tolist()[0]
    LIPICON_call = doctor_call.LIPICON.tolist()[0]
    AGLIP_call = doctor_call.AGLIP.tolist()[0]
    CIFIBET_call = doctor_call.CIFIBET.tolist()[0]
    AMLEVO_call = doctor_call.AMLEVO.tolist()[0]
    CARDOBIS_call = doctor_call.CARDOBIS.tolist()[0]
    RIVAROX_call = doctor_call.RIVAROX.tolist()[0]
    NOCLOG_call = doctor_call.NOCLOG.tolist()[0]
    BEMPID_call = doctor_call.BEMPID.tolist()[0]
    AROTIDE_call = doctor_call.AROTIDE.tolist()[0]
    FOBUNID_call = doctor_call.FOBUNID.tolist()[0]

    img_draw.text((150, 375), '(' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((40, 450), number_decorator(LIGZID_call), (0, 0, 0), font1)
    img_draw.text((135, 450), number_decorator(EMAZID_call), (0, 0, 0), font1)
    img_draw.text((235, 450), number_decorator(LIPICON_call), (0, 0, 0), font1)
    img_draw.text((335, 450), number_decorator(AGLIP_call), (0, 0, 0), font1)
    img_draw.text((435, 450), number_decorator(CIFIBET_call), (0, 0, 0), font1)
    img_draw.text((535, 450), number_decorator(AMLEVO_call), (0, 0, 0), font1)
    img_draw.text((635, 450), number_decorator(CARDOBIS_call), (0, 0, 0), font1)
    img_draw.text((735, 450), number_decorator(RIVAROX_call), (0, 0, 0), font1)
    img_draw.text((835, 450), number_decorator(NOCLOG_call), (0, 0, 0), font1)
    img_draw.text((935, 450), number_decorator(BEMPID_call), (0, 0, 0), font1)
    img_draw.text((1035, 450), number_decorator(AROTIDE_call), (0, 0, 0), font1)
    img_draw.text((1135, 450), number_decorator(FOBUNID_call), (0, 0, 0), font1)
    img1.save("./Images/all_kpi_image.png")

    print("3. All KPI Image created")


def rsm_kpi_image(rsm):
    img1 = Image.open("./Images/kpi.png")
    img_draw = ImageDraw.Draw(img1)

    img1 = Image.open("./Images/new_file.png")
    img_draw = ImageDraw.Draw(img1)

    sales_achiv = pd.read_excel("./Data/SalesAchivment/SalesAchivment_"+str(rsm)+".xlsx").iloc[0]

    LIGZID_achiv = sales_achiv.LIGAZID
    EMAZID_achiv = sales_achiv.EMAZID
    LIPICON_achiv = sales_achiv.LIPICON
    AGLIP_achiv = sales_achiv.AGLIP
    CIFIBET_achiv = sales_achiv.CIFIBET
    AMLEVO_achiv = sales_achiv.AMLEVO
    CARDOBIS_achiv = sales_achiv.CARDOBIS
    RIVAROX_achiv = sales_achiv.RIVAROX
    NOCLOG_achiv = sales_achiv.NOCLOG
    BEMPID_achiv = sales_achiv.BEMPID
    AROTIDE_achiv = sales_achiv.AROTIDE
    FOBUNID_achiv = sales_achiv.FOBUNID

    img_draw.text((265, 13), 'up to (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((25, 85), number_decorator(LIGZID_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((120, 85), number_decorator(EMAZID_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((217, 85), number_decorator(LIPICON_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((320, 85), number_decorator(AGLIP_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((417, 85), number_decorator(CIFIBET_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((517, 85), number_decorator(AMLEVO_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((617, 85), number_decorator(CARDOBIS_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((717, 85), number_decorator(RIVAROX_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((817, 85), number_decorator(NOCLOG_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((917, 85), number_decorator(BEMPID_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((1017, 85), number_decorator(AROTIDE_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((1117, 85), number_decorator(FOBUNID_achiv) + '%', (0, 0, 0), font1)

    # # ------------------------------------------------------------------------------
    sales_trend = pd.read_excel("./Data/SalesTrend/SalesTrend_"+str(rsm)+".xlsx").iloc[0]
    LIGZID_trend = sales_trend.LIGAZID
    EMAZID_trend = sales_trend.EMAZID
    LIPICON_trend = sales_trend.LIPICON
    AGLIP_trend = sales_trend.AGLIP
    CIFIBET_trend = sales_trend.CIFIBET
    AMLEVO_trend = sales_trend.AMLEVO
    CARDOBIS_trend = sales_trend.CARDOBIS
    RIVAROX_trend = sales_trend.RIVAROX
    NOCLOG_trend = sales_trend.NOCLOG
    BEMPID_trend = sales_trend.BEMPID
    AROTIDE_trend = sales_trend.AROTIDE
    FOBUNID_trend = sales_trend.FOBUNID

    img_draw.text((180, 130), 'up to (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((25, 200), number_decorator(LIGZID_trend) + '%', (0, 0, 0), font1)
    img_draw.text((120, 200), number_decorator(EMAZID_trend) + '%', (0, 0, 0), font1)
    img_draw.text((217, 200), number_decorator(LIPICON_trend) + '%', (0, 0, 0), font1)
    img_draw.text((320, 200), number_decorator(AGLIP_trend) + '%', (0, 0, 0), font1)
    img_draw.text((417, 200), number_decorator(CIFIBET_trend) + '%', (0, 0, 0), font1)
    img_draw.text((517, 200), number_decorator(AMLEVO_trend) + '%', (0, 0, 0), font1)
    img_draw.text((617, 200), number_decorator(CARDOBIS_trend) + '%', (0, 0, 0), font1)
    img_draw.text((717, 200), number_decorator(RIVAROX_trend) + '%', (0, 0, 0), font1)
    img_draw.text((817, 200), number_decorator(NOCLOG_trend) + '%', (0, 0, 0), font1)
    img_draw.text((917, 200), number_decorator(BEMPID_trend) + '%', (0, 0, 0), font1)
    img_draw.text((1017, 200), number_decorator(AROTIDE_trend) + '%', (0, 0, 0), font1)
    img_draw.text((1117, 200), number_decorator(FOBUNID_trend) + '%', (0, 0, 0), font1)

    # # # ------------------------------ Seen Rx ----------------------------------------
    seen_rx = pd.read_excel('./Data/SeenRx/Seen_Rx_Data.xlsx')
    LIGZID_rx = seen_rx.LIGAZID.tolist()[0]
    EMAZID_rx = seen_rx.EMAZID.tolist()[0]
    LIPICON_rx = seen_rx.LIPICON.tolist()[0]
    AGLIP_rx = seen_rx.AGLIP.tolist()[0]
    CIFIBET_rx = seen_rx.CIFIBET.tolist()[0]
    AMLEVO_rx = seen_rx.AMLEVO.tolist()[0]
    CARDOBIS_rx = seen_rx.CARDOBIS.tolist()[0]
    RIVAROX_rx = seen_rx.RIVAROX.tolist()[0]
    NOCLOG_rx = seen_rx.NOCLOG.tolist()[0]
    BEMPID_rx = seen_rx.BEMPID.tolist()[0]
    AROTIDE_rx = seen_rx.AROTIDE.tolist()[0]
    FOBUNID_rx = seen_rx.FOBUNID.tolist()[0]

    img_draw.text((120, 247), '(' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((40, 325), number_decorator(LIGZID_rx), (0, 0, 0), font1)
    img_draw.text((135, 325), number_decorator(EMAZID_rx), (0, 0, 0), font1)
    img_draw.text((235, 325), number_decorator(LIPICON_rx), (0, 0, 0), font1)
    img_draw.text((335, 325), number_decorator(AGLIP_rx), (0, 0, 0), font1)
    img_draw.text((435, 325), number_decorator(CIFIBET_rx), (0, 0, 0), font1)
    img_draw.text((535, 325), number_decorator(AMLEVO_rx), (0, 0, 0), font1)
    img_draw.text((635, 325), number_decorator(CARDOBIS_rx), (0, 0, 0), font1)
    img_draw.text((735, 325), number_decorator(RIVAROX_rx), (0, 0, 0), font1)
    img_draw.text((835, 325), number_decorator(NOCLOG_rx), (0, 0, 0), font1)
    img_draw.text((935, 325), number_decorator(BEMPID_rx), (0, 0, 0), font1)
    img_draw.text((1035, 325), number_decorator(AROTIDE_rx), (0, 0, 0), font1)
    img_draw.text((1135, 325), number_decorator(FOBUNID_rx), (0, 0, 0), font1)

    # # ------------------------------ Doctor Call -----------------------------------
    doctor_call = pd.read_excel('./Data/Call/doctor_call_data.xlsx')
    LIGZID_call = doctor_call.LIGAZID.tolist()[0]
    EMAZID_call = doctor_call.EMAZID.tolist()[0]
    LIPICON_call = doctor_call.LIPICON.tolist()[0]
    AGLIP_call = doctor_call.AGLIP.tolist()[0]
    CIFIBET_call = doctor_call.CIFIBET.tolist()[0]
    AMLEVO_call = doctor_call.AMLEVO.tolist()[0]
    CARDOBIS_call = doctor_call.CARDOBIS.tolist()[0]
    RIVAROX_call = doctor_call.RIVAROX.tolist()[0]
    NOCLOG_call = doctor_call.NOCLOG.tolist()[0]
    BEMPID_call = doctor_call.BEMPID.tolist()[0]
    AROTIDE_call = doctor_call.AROTIDE.tolist()[0]
    FOBUNID_call = doctor_call.FOBUNID.tolist()[0]

    img_draw.text((150, 375), '(' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((40, 450), number_decorator(LIGZID_call), (0, 0, 0), font1)
    img_draw.text((135, 450), number_decorator(EMAZID_call), (0, 0, 0), font1)
    img_draw.text((235, 450), number_decorator(LIPICON_call), (0, 0, 0), font1)
    img_draw.text((335, 450), number_decorator(AGLIP_call), (0, 0, 0), font1)
    img_draw.text((435, 450), number_decorator(CIFIBET_call), (0, 0, 0), font1)
    img_draw.text((535, 450), number_decorator(AMLEVO_call), (0, 0, 0), font1)
    img_draw.text((635, 450), number_decorator(CARDOBIS_call), (0, 0, 0), font1)
    img_draw.text((740, 450), number_decorator(RIVAROX_call), (0, 0, 0), font1)
    img_draw.text((835, 450), number_decorator(NOCLOG_call), (0, 0, 0), font1)
    img_draw.text((935, 450), number_decorator(BEMPID_call), (0, 0, 0), font1)
    img_draw.text((1035, 450), number_decorator(AROTIDE_call), (0, 0, 0), font1)
    img_draw.text((1135, 450), number_decorator(FOBUNID_call), (0, 0, 0), font1)
    img1.save("./Images/kpi_image_" + str(rsm) + ".png")
    print('3. Speceric KPI Image Generated')

