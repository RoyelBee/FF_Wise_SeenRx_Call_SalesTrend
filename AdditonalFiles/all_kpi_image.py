import pandas as pd
from PIL import ImageDraw, ImageFont
from PIL import Image
import SQLConnection as conn
from datetime import datetime
import Path as dir
from datetime import date, timedelta
import warnings

warnings.filterwarnings('ignore')

today = datetime.today()
yesterday = date.today() - timedelta(days=1)

day = str(yesterday.day) + '-' + str(yesterday.strftime("%b")) + '-' + str(yesterday.year)

curr_month = str(today.strftime("%b")) + '-' + str(today.year)

font = ImageFont.truetype("./font/ROCK.ttf", 30, encoding="unic")
font1 = ImageFont.truetype("./font/ROCK.ttf", 30, encoding="unic")


def number_decorator(number):
    number = round(number, 2)
    number = format(number, ',')
    return number


def all_kpi_images():
    img1 = Image.open("./Images/kpi_image.png")
    img_draw = ImageDraw.Draw(img1)

    sales_query = (""" DECLARE @YearMonth as VARCHAR(6) = Convert (varchar,Getdate(),112)
        DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), GETDATE(), 112)
        DECLARE @FIRSTDATEOFMONTH AS CHAR(8) = CONVERT(VARCHAR(6), GETDATE(), 112)+'01'
        DECLARE @Today as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)
        DECLARE @LastDay as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)-1
        DECLARE @TotalDaysInMonth as Integer=(SELECT DATEDIFF(DAY, getdate(), DATEADD(MONTH, 1, getdate())))
        DECLARE @TotalDaysGone as integer =(select  count(distinct transdate) from OESalesSummery where left(transdate,6)=@This_month and  TRANSDATE<=@LastDay)

        Select
        isnull(Sum(Case when FFTarget.Brand = 'LIGAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS LIGAZID,
        isnull(Sum(Case when FFTarget.Brand = 'EMAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS EMAZID,
        isnull(Sum(Case when FFTarget.Brand = 'LIPICON' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS LIPICON,
        isnull(Sum(Case when FFTarget.Brand = 'AGLIP' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS AGLIP,
        isnull(Sum(Case when FFTarget.Brand = 'CIFIBET' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS CIFIBET,
        isnull(Sum(Case when FFTarget.Brand = 'AMLEVO' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS AMLEVO,
        isnull(Sum(Case when FFTarget.Brand = 'CARDOBIS' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS CARDOBIS,
        isnull(Sum(Case when FFTarget.Brand = 'RIVAROX' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS RIVAROX,
        isnull(Sum(Case when FFTarget.Brand = 'NOCLOG' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS NOCLOG

        from
        (Select BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_Target group by BRAND) as FFTarget
         left join
        (Select BRAND, SUM(extinvmisc) AS SalesAmount  from V_FF_Brand_Sales where TRANSDATE <= @LastDay  group by BRAND) as FFSales
        ON  (FFTarget.BRAND=FFSales.BRAND)
    """)
    sales_data = pd.read_sql_query(sales_query, conn.db_connection())

    LIGZID_data = sales_data.LIGAZID.tolist()[0]
    EMAZID_data = sales_data.EMAZID.tolist()[0]
    LIPICON_data = sales_data.LIPICON.tolist()[0]
    AGLIP_data = sales_data.AGLIP.tolist()[0]
    CIFIBET_data = sales_data.CIFIBET.tolist()[0]
    AMLEVO_data = sales_data.AMLEVO.tolist()[0]
    CARDOBIS_data = sales_data.CARDOBIS.tolist()[0]
    RIVAROX_data = sales_data.RIVAROX.tolist()[0]
    NOCLOG_data = sales_data.NOCLOG.tolist()[0]

    # target_tag.text((500, 59), number_decorator(AGLIP_data) + '%', (0, 0, 0), font)
    img_draw.text((8, 15), 'Sales Trend % (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((25, 100), number_decorator(LIGZID_data) + '%', (0, 0, 0), font1)
    img_draw.text((170, 100), number_decorator(EMAZID_data) + '%', (0, 0, 0), font1)
    img_draw.text((310, 100), number_decorator(LIPICON_data) + '%', (0, 0, 0), font1)
    img_draw.text((455, 100), number_decorator(AGLIP_data) + '%', (0, 0, 0), font1)
    img_draw.text((600, 100), number_decorator(CIFIBET_data) + '%', (0, 0, 0), font1)
    img_draw.text((740, 100), number_decorator(AMLEVO_data) + '%', (0, 0, 0), font1)
    img_draw.text((885, 100), number_decorator(CARDOBIS_data) + '%', (0, 0, 0), font1)
    img_draw.text((1020, 100), number_decorator(RIVAROX_data) + '%', (0, 0, 0), font1)
    img_draw.text((1160, 100), number_decorator(NOCLOG_data) + '%', (0, 0, 0), font1)

    seenrx_query = ("""
            Select
            sum([LIGAZID Seen Rx]) as LIGZID,
            sum([EMAZID Seen Rx]) as EMAZID,
            sum([LIPICON Seen Rx]) as LIPICON,
            sum([AGLIP Seen Rx]) as AGLIP,
            sum([CIFIBET Seen Rx]) as CIFIBET,
            sum([AMLEVO Seen Rx]) as AMLEVO,
            sum([CARDOBIS Seen Rx]) as CARDOBIS,
            sum([RIVAROX Seen Rx]) as RIVAROX,
            sum([NOCLOG Seen Rx]) as NOCLOG
            from v_LastDay_SeenRx where ID = 2
            """)

    seenrx_seen_data = pd.read_sql_query(seenrx_query, conn.db_connection_dcr())

    LIGZID_seen_seen_data = seenrx_seen_data.LIGZID.tolist()[0]
    EMAZID_seen_data = seenrx_seen_data.EMAZID.tolist()[0]
    LIPICON_seen_data = seenrx_seen_data.LIPICON.tolist()[0]
    AGLIP_seen_data = seenrx_seen_data.AGLIP.tolist()[0]
    CIFIBET_seen_data = seenrx_seen_data.CIFIBET.tolist()[0]
    AMLEVO_seen_data = seenrx_seen_data.AMLEVO.tolist()[0]
    CARDOBIS_seen_data = seenrx_seen_data.CARDOBIS.tolist()[0]
    RIVAROX_seen_data = seenrx_seen_data.RIVAROX.tolist()[0]
    NOCLOG_seen_data = seenrx_seen_data.NOCLOG.tolist()[0]

    # # Seen Rx
    img_draw.text((8, 200), 'SeenRx (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((45, 280), number_decorator(LIGZID_seen_seen_data), (0, 0, 0), font1)
    img_draw.text((190, 280), number_decorator(EMAZID_seen_data), (0, 0, 0), font1)
    img_draw.text((330, 280), number_decorator(LIPICON_seen_data), (0, 0, 0), font1)
    img_draw.text((475, 280), number_decorator(AGLIP_seen_data), (0, 0, 0), font1)
    img_draw.text((620, 280), number_decorator(CIFIBET_seen_data), (0, 0, 0), font1)
    img_draw.text((760, 280), number_decorator(AMLEVO_seen_data), (0, 0, 0), font1)
    img_draw.text((910, 280), number_decorator(CARDOBIS_seen_data), (0, 0, 0), font1)
    img_draw.text((1040, 280), number_decorator(RIVAROX_seen_data), (0, 0, 0), font1)
    img_draw.text((1180, 280), number_decorator(NOCLOG_seen_data), (0, 0, 0), font1)

    call_data = pd.read_excel('./Data/Call/doctor_call_data.xlsx')
    LIGAZID_Call = sum(call_data['LIGAZID Call'])
    EMAZID_Call = sum(call_data['EMAZID Call'])
    LIPICON_Call = sum(call_data['LIPICON Call'])
    AGLIP_Call = sum(call_data['AGLIP Call'])
    CIFIBET_Call = sum(call_data['CIFIBET Call'])
    AMLEVO_Call = sum(call_data['AMLEVO Call'])
    CARDOBIS_Call = sum(call_data['CARDOBIS Call'])
    RIVAROX_Call = sum(call_data['RIVAROX Call'])
    Noclog_Call = sum(call_data['Noclog Call'])

    # # Doctor Call
    img_draw.text((8, 380), 'Doctor Call (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((25, 460), number_decorator(LIGAZID_Call), (0, 0, 0), font1)
    img_draw.text((170, 460), number_decorator(EMAZID_Call), (0, 0, 0), font1)
    img_draw.text((310, 460), number_decorator(LIPICON_Call), (0, 0, 0), font1)
    img_draw.text((455, 460), number_decorator(AGLIP_Call), (0, 0, 0), font1)
    img_draw.text((600, 460), number_decorator(CIFIBET_Call), (0, 0, 0), font1)
    img_draw.text((740, 460), number_decorator(AMLEVO_Call), (0, 0, 0), font1)
    img_draw.text((885, 460), number_decorator(CARDOBIS_Call), (0, 0, 0), font1)
    img_draw.text((1020, 460), number_decorator(RIVAROX_Call), (0, 0, 0), font1)
    img_draw.text((1160, 460), number_decorator(Noclog_Call), (0, 0, 0), font1)
    img1.save("./Images/all_kpi_image.png")

    print("5.2. All KPI Image created")


def kpi_image(rsm):
    img1 = Image.open("./Images/kpi_image.png")
    img_draw = ImageDraw.Draw(img1)

    # sales_query = (""" DECLARE @YearMonth as VARCHAR(6) = Convert (varchar,Getdate(),112)
    #     DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), GETDATE(), 112)
    #     DECLARE @FIRSTDATEOFMONTH AS CHAR(8) = CONVERT(VARCHAR(6), GETDATE(), 112)+'01'
    #     DECLARE @Today as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)
    #     DECLARE @LastDay as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)-1
    #     DECLARE @TotalDaysInMonth as Integer=(SELECT DATEDIFF(DAY, getdate(), DATEADD(MONTH, 1, getdate())))
    #     DECLARE @TotalDaysGone as integer =(select  count(distinct transdate) from OESalesSummery where left(transdate,6)=@This_month and  TRANSDATE<=@LastDay)
    # 
    #     Select
    #     isnull(Sum(Case when FFTarget.Brand = 'LIGAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS LIGAZID,
    #     isnull(Sum(Case when FFTarget.Brand = 'EMAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS EMAZID,
    #     isnull(Sum(Case when FFTarget.Brand = 'LIPICON' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS LIPICON,
    #     isnull(Sum(Case when FFTarget.Brand = 'AGLIP' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS AGLIP,
    #     isnull(Sum(Case when FFTarget.Brand = 'CIFIBET' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS CIFIBET,
    #     isnull(Sum(Case when FFTarget.Brand = 'AMLEVO' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS AMLEVO,
    #     isnull(Sum(Case when FFTarget.Brand = 'CARDOBIS' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS CARDOBIS,
    #     isnull(Sum(Case when FFTarget.Brand = 'RIVAROX' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS RIVAROX,
    #     isnull(Sum(Case when FFTarget.Brand = 'NOCLOG' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS NOCLOG
    # 
    #     from
    #     (Select BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_Target group by BRAND) as FFTarget
    #      left join
    #     (Select BRAND, SUM(extinvmisc) AS SalesAmount  from V_FF_Brand_Sales where TRANSDATE <= @LastDay  group by BRAND) as FFSales
    #     ON  (FFTarget.BRAND=FFSales.BRAND)
    # """)
    # sales_data = pd.read_sql_query(sales_query, conn.db_connection())
    sales_data = pd.read_excel("./Data/SalesTrend/SalesTrend_" + str(rsm) + ".xlsx")
    LIGZID_data = sales_data.LIGAZID.tolist()[0]
    EMAZID_data = sales_data.EMAZID.tolist()[0]
    LIPICON_data = sales_data.LIPICON.tolist()[0]
    AGLIP_data = sales_data.AGLIP.tolist()[0]
    CIFIBET_data = sales_data.CIFIBET.tolist()[0]
    AMLEVO_data = sales_data.AMLEVO.tolist()[0]
    CARDOBIS_data = sales_data.CARDOBIS.tolist()[0]
    RIVAROX_data = sales_data.RIVAROX.tolist()[0]
    NOCLOG_data = sales_data.NOCLOG.tolist()[0]

    img_draw.text((8, 15), 'Sales Trend % (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((25, 100), number_decorator(LIGZID_data) + '%', (0, 0, 0), font1)
    img_draw.text((170, 100), number_decorator(EMAZID_data) + '%', (0, 0, 0), font1)
    img_draw.text((310, 100), number_decorator(LIPICON_data) + '%', (0, 0, 0), font1)
    img_draw.text((455, 100), number_decorator(AGLIP_data) + '%', (0, 0, 0), font1)
    img_draw.text((600, 100), number_decorator(CIFIBET_data) + '%', (0, 0, 0), font1)
    img_draw.text((740, 100), number_decorator(AMLEVO_data) + '%', (0, 0, 0), font1)
    img_draw.text((885, 100), number_decorator(CARDOBIS_data) + '%', (0, 0, 0), font1)
    img_draw.text((1020, 100), number_decorator(RIVAROX_data) + '%', (0, 0, 0), font1)
    img_draw.text((1160, 100), number_decorator(NOCLOG_data) + '%', (0, 0, 0), font1)

    seenrx_query = (""" 
                select * from
                (Select
                left([FF ID],3) as [FF ID],
                isnull(sum([LIGAZID Seen Rx]),0)  as [LIGAZID],
                isnull(sum([EMAZID Seen Rx]),0) as [EMAZID],
                isnull(sum([LIPICON Seen Rx]),0) as [LIPICON],
                isnull(sum([AGLIP Seen Rx]),0) as [AGLIP],
                isnull(sum([CIFIBET Seen Rx]),0) as [CIFIBET],
                isnull(sum([AMLEVO Seen Rx]),0) as [AMLEVO],
                isnull(sum([CARDOBIS Seen Rx]),0) as [CARDOBIS],
                isnull(sum([RIVAROX Seen Rx]),0) as [RIVAROX],
                isnull(sum([NOCLOG Seen Rx]),0) as [NOCLOG]
                from v_LastDay_SeenRx 
                where [FF ID]  = left([FF ID],4)+'0' and left([FF ID],3) like ?
                group by left([FF ID],3)
                
                union all
                
                Select [FF ID],
                isnull([LIGAZID Seen Rx],0) as [LIGAZID Seen Rx],
                isnull([EMAZID Seen Rx],0) as [EMAZID Seen Rx],
                ISNULL([LIPICON Seen Rx], 0) as [LIPICON Seen Rx],
                ISNULL([AGLIP Seen Rx], 0) as [AGLIP Seen Rx],
                ISNULL([CIFIBET Seen Rx], 0) as [CIFIBET Seen Rx],
                ISNULL([AMLEVO Seen Rx], 0) as [AMLEVO Seen Rx],
                ISNULL([CARDOBIS Seen Rx], 0) as [CARDOBIS Seen Rx],
                ISNULL([RIVAROX Seen Rx], 0) as [RIVAROX Seen Rx],
                ISNULL([NOCLOG Seen Rx], 0) as [NOCLOG Seen Rx]
                from v_LastDay_SeenRx where  left([FF ID],3) like ?
                
                ) as T1
                order by [FF ID] asc
            """)

    seenrx_seen_data = pd.read_sql_query(seenrx_query, conn.db_connection_dcr(), params=(rsm, rsm))

    LIGZID_seen_seen_data = seenrx_seen_data.LIGAZID.tolist()[0]
    EMAZID_seen_data = seenrx_seen_data.EMAZID.tolist()[0]
    LIPICON_seen_data = seenrx_seen_data.LIPICON.tolist()[0]
    AGLIP_seen_data = seenrx_seen_data.AGLIP.tolist()[0]
    CIFIBET_seen_data = seenrx_seen_data.CIFIBET.tolist()[0]
    AMLEVO_seen_data = seenrx_seen_data.AMLEVO.tolist()[0]
    CARDOBIS_seen_data = seenrx_seen_data.CARDOBIS.tolist()[0]
    RIVAROX_seen_data = seenrx_seen_data.RIVAROX.tolist()[0]
    NOCLOG_seen_data = seenrx_seen_data.NOCLOG.tolist()[0]

    # # Seen Rx
    img_draw.text((8, 200), 'SeenRx (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((45, 280), number_decorator(LIGZID_seen_seen_data), (0, 0, 0), font1)
    img_draw.text((190, 280), number_decorator(EMAZID_seen_data), (0, 0, 0), font1)
    img_draw.text((330, 280), number_decorator(LIPICON_seen_data), (0, 0, 0), font1)
    img_draw.text((475, 280), number_decorator(AGLIP_seen_data), (0, 0, 0), font1)
    img_draw.text((620, 280), number_decorator(CIFIBET_seen_data), (0, 0, 0), font1)
    img_draw.text((760, 280), number_decorator(AMLEVO_seen_data), (0, 0, 0), font1)
    img_draw.text((910, 280), number_decorator(CARDOBIS_seen_data), (0, 0, 0), font1)
    img_draw.text((1040, 280), number_decorator(RIVAROX_seen_data), (0, 0, 0), font1)
    img_draw.text((1180, 280), number_decorator(NOCLOG_seen_data), (0, 0, 0), font1)

    # call_data = pd.read_excel('./Data/Call/doctor_call_data.xlsx')
    # LIGAZID_Call = sum(call_data['LIGAZID Call'])
    # EMAZID_Call = sum(call_data['EMAZID Call'])
    # LIPICON_Call = sum(call_data['LIPICON Call'])
    # AGLIP_Call = sum(call_data['AGLIP Call'])
    # CIFIBET_Call = sum(call_data['CIFIBET Call'])
    # AMLEVO_Call = sum(call_data['AMLEVO Call'])
    # CARDOBIS_Call = sum(call_data['CARDOBIS Call'])
    # RIVAROX_Call = sum(call_data['RIVAROX Call'])
    # Noclog_Call = sum(call_data['Noclog Call'])
    #
    # # # Doctor Call
    # img_draw.text((8, 380), 'Doctor Call (' + str(day) + ')', (255, 255, 255), font)
    # img_draw.text((25, 460), number_decorator(LIGAZID_Call), (0, 0, 0), font1)
    # img_draw.text((170, 460), number_decorator(EMAZID_Call), (0, 0, 0), font1)
    # img_draw.text((310, 460), number_decorator(LIPICON_Call), (0, 0, 0), font1)
    # img_draw.text((455, 460), number_decorator(AGLIP_Call), (0, 0, 0), font1)
    # img_draw.text((600, 460), number_decorator(CIFIBET_Call), (0, 0, 0), font1)
    # img_draw.text((740, 460), number_decorator(AMLEVO_Call), (0, 0, 0), font1)
    # img_draw.text((885, 460), number_decorator(CARDOBIS_Call), (0, 0, 0), font1)
    # img_draw.text((1020, 460), number_decorator(RIVAROX_Call), (0, 0, 0), font1)
    # img_draw.text((1160, 460), number_decorator(Noclog_Call), (0, 0, 0), font1)
    img1.save("./Images/kpi_image_" + str(rsm) + ".png")

    print("5.2. All KPI Image created")
