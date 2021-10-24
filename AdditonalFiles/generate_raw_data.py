import pandas as pd
import db_connection as conn


def seen_rx_data():
    seen_rx_data = pd.read_sql_query(""" SELECT [FFTR], [LIGAZID Seen Rx], [EMAZID Seen Rx], [LIPICON Seen Rx],
    [AGLIP Seen Rx], [CIFIBET Seen Rx], [AMLEVO Seen Rx], [CARDOBIS Seen Rx], [RIVAROX Seen Rx],[Noclog Seen Rx]

    FROM (SELECT LEFT(area_id, 4) + '0' AS FFTR,
        COUNT(CASE WHEN medicine_id = 'LIGAZID' THEN doctor_id END) AS [LIGAZID Seen Rx],
        COUNT(CASE WHEN medicine_id = 'EMAZID' THEN doctor_id END) AS [EMAZID Seen Rx],
        COUNT(CASE WHEN medicine_id = 'LIPICON' THEN doctor_id END) AS [LIPICON Seen Rx],
        COUNT(CASE WHEN medicine_id = 'AGLIP' THEN doctor_id END) AS [AGLIP Seen Rx],
        COUNT(CASE WHEN medicine_id = 'CIFIBET' THEN doctor_id END) AS [CIFIBET Seen Rx],
        COUNT(CASE WHEN medicine_id = 'AMLEVO' THEN doctor_id END) AS [AMLEVO Seen Rx],
        COUNT(CASE WHEN medicine_id = 'CARDOBIS' THEN doctor_id END) AS [CARDOBIS Seen Rx],
        COUNT(CASE WHEN medicine_id = 'RIVAROX' THEN doctor_id END) AS [RIVAROX Seen Rx],
        COUNT(CASE WHEN medicine_id = 'Noclog' THEN doctor_id END) AS [Noclog Seen Rx]
        FROM  dbo.sm_prescription_seen_details
        WHERE   (submit_date = DATEADD(day, - 1, CAST(GETDATE() AS date))) AND (LEFT(area_id, 1) = 'C')
        GROUP BY LEFT(area_id, 4) + '0'
        UNION ALL
        SELECT area_id,
            COUNT(CASE WHEN medicine_id = 'LIGAZID' THEN doctor_id END) AS [LIGAZID Seen Rx],
            COUNT(CASE WHEN medicine_id = 'EMAZID' THEN doctor_id END) AS [EMAZID Seen Rx],
            COUNT(CASE WHEN medicine_id = 'LIPICON' THEN doctor_id END) AS [LIPICON Seen Rx],
            COUNT(CASE WHEN medicine_id = 'AGLIP' THEN doctor_id END) AS [AGLIP Seen Rx],
            COUNT(CASE WHEN medicine_id = 'CIFIBET' THEN doctor_id END) AS [CIFIBET Seen Rx],
            COUNT(CASE WHEN medicine_id = 'AMLEVO' THEN doctor_id END) AS [AMLEVO Seen Rx],
            COUNT(CASE WHEN medicine_id = 'CARDOBIS' THEN doctor_id END) AS [CARDOBIS Seen Rx],
            COUNT(CASE WHEN medicine_id = 'RIVAROX' THEN doctor_id END) AS [RIVAROX Seen Rx],
            COUNT(CASE WHEN medicine_id = 'Noclog' THEN doctor_id END) AS [Noclog Seen Rx]
        FROM     dbo.sm_prescription_seen_details AS sm_prescription_seen_details_1
        WHERE  (submit_date = DATEADD(day, - 1, CAST(GETDATE() AS date))) AND (LEFT(area_id, 1) = 'C')
        GROUP BY area_id) AS FF_SeenRx """, conn.m_reporting)
    seen_rx_data.to_excel("./Data/SeenRx/Seen_Rx_Data.xlsx", index=False)

    data = pd.read_excel('./Data/SeenRx/Seen_Rx_Data.xlsx')

    CBU = data[data['FFTR'].str.contains('CBU')]
    CBU.to_excel('./Data/SeenRx/SeenRx_CBU.xlsx', index=False)

    CCF = data[data['FFTR'].str.contains('CCF')]
    CCF.to_excel('./Data/SeenRx/SeenRx_CCF.xlsx', index=False)

    CCX = data[data['FFTR'].str.contains('CCX')]
    CCX.to_excel('./Data/SeenRx/SeenRx_CCX.xlsx', index=False)

    CNH = data[data['FFTR'].str.contains('CNH')]
    CNH.to_excel('./Data/SeenRx/SeenRx_CNH.xlsx', index=False)

    CKJ = data[data['FFTR'].str.contains('CKJ')]
    CKJ.to_excel('./Data/SeenRx/SeenRx_CKJ.xlsx', index=False)

    CMT = data[data['FFTR'].str.contains('CMT')]
    CMT.to_excel('./Data/SeenRx/SeenRx_CMT.xlsx', index=False)

    CRB = data[data['FFTR'].str.contains('CRB')]
    CRB.to_excel('./Data/SeenRx/SeenRx_CRB.xlsx', index=False)

    CRP = data[data['FFTR'].str.contains('CRP')]
    CRP.to_excel('./Data/SeenRx_CRP.xlsx', index=False)

    CSB = data[data['FFTR'].str.contains('CSB')]
    CSB.to_excel('./Data/SeenRx/SeenRx_CSB.xlsx', index=False)
    print('1. Seen Rx Data Saved')


def doctor_call_data():
    doctor_call_data = pd.read_sql_query(""" SELECT  RTRIM(route_id) AS FFTR,
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'LIGAZID' THEN visit_sl END) AS [LIGAZID Call],
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'EMAZID' THEN visit_sl END) AS [EMAZID Call],
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'LIPICON' THEN visit_sl END) AS [LIPICON Call],
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'AGLIP' THEN visit_sl END) AS [AGLIP Call],
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'CIFIBET' THEN visit_sl END) AS [CIFIBET Call],
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'AMLEVO' THEN visit_sl END) AS [AMLEVO Call],
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'CARDOBIS' THEN visit_sl END) AS [CARDOBIS Call],
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'RIVAROX' THEN visit_sl END) AS [RIVAROX Call],
    COUNT( distinct CASE WHEN RTRIM(BRAND) = 'Noclog' THEN visit_sl END) AS [Noclog Call]
    FROM    dbo.drBrandEffectiveCall
    WHERE    visit_date = DATEADD(day, - 1, CAST(GETDATE() AS date))
    group by RTRIM(route_id)
    order by FFTR asc """, conn.m_reporting)
    doctor_call_data.to_excel("./Data/Call/doctor_call_data.xlsx", index=False)

    data = pd.read_excel('./Data/Call/doctor_call_data.xlsx')

    CBU = data[data['FFTR'].str.contains('CBU')]
    CBU.to_excel('./Data/Call/Call_CBU.xlsx', index=False)

    CCF = data[data['FFTR'].str.contains('CCF')]
    CCF.to_excel('./Data/Call/Call_CCF.xlsx', index=False)

    CCX = data[data['FFTR'].str.contains('CCX')]
    CCX.to_excel('./Data/Call/Call_CCX.xlsx', index=False)

    CNH = data[data['FFTR'].str.contains('CNH')]
    CNH.to_excel('./Data/Call/Call_CNH.xlsx', index=False)

    CKJ = data[data['FFTR'].str.contains('CKJ')]
    CKJ.to_excel('./Data/Call/Call_CKJ.xlsx', index=False)

    CMT = data[data['FFTR'].str.contains('CMT')]
    CMT.to_excel('./Data/Call/Call_CMT.xlsx', index=False)

    CRB = data[data['FFTR'].str.contains('CRB')]
    CRB.to_excel('./Data/Call/Call_CRB.xlsx', index=False)

    CRP = data[data['FFTR'].str.contains('CRP')]
    CRP.to_excel('./Data/Call_CRP.xlsx', index=False)

    CSB = data[data['FFTR'].str.contains('CSB')]
    CSB.to_excel('./Data/Call/Call_CSB.xlsx', index=False)

    print('2. Doctor Call Data Saved')


def sales_trend_data():
    sales_trend_data = pd.read_sql_query(""" DECLARE @YearMonth as VARCHAR(6) = Convert (varchar,Getdate()-1,112)
    DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), GETDATE()-1, 112)
    DECLARE @FIRSTDATEOFMONTH AS CHAR(8) = CONVERT(VARCHAR(6), GETDATE()-1, 112)+'01'
    DECLARE @Today as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112) 
    DECLARE @LastDay as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)-1 
    DECLARE @TotalDaysInMonth as Integer=(SELECT DATEDIFF(DAY, getdate()-1, DATEADD(MONTH, 1, getdate())))
    DECLARE @TotalDaysGone as integer =(select  count(distinct transdate) from OESalesSummery where left(transdate,6)=@This_month and  TRANSDATE<=@LastDay)
    
    Select * from 
    (Select 
    FFTarget.FFTR,
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'LIGAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [LIGAZID],
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'EMAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [EMAZID],
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'LIPICON' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [LIPICON],
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'AGLIP' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [AGLIP],
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'CIFIBET' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [CIFIBET],
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'AMLEVO' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [AMLEVO],
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'CARDOBIS' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [CARDOBIS],
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Rivarox' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [RIVAROX],
    Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'NOCLOG' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [NOCLOG]
    
    from 
    (
    Select 'RSM' AS FF, RSMTR as FFTR,BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_Target 
    group by RSMTR,BRAND
    union all
    Select 'FM' AS FF, FMTR as FFTR,BRAND,SUM(Amount) AS TargetAmount  from V_FF_Brand_Target 
    group by FMTR,BRAND
    union all
    Select 'MSO' AS FF, MSOTR as FFTR,BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_Target 
    group by MSOTR,BRAND
    ) as FFTarget 
    left join
    (
    Select 'RSM' AS FF, left(MSOTR,3)  as FFTR,BRAND, SUM(extinvmisc) AS SalesAmount  from V_FF_Brand_Sales 
    where TRANSDATE <= @LastDay
    group by left(MSOTR,3) ,BRAND
    union all
    Select 'FM' AS FF, left(MSOTR,4)+'0'  as FFTR,BRAND,SUM(extinvmisc) AS SalesAmount   from V_FF_Brand_Sales 
    where TRANSDATE <= @LastDay
    group by left(MSOTR,4)+'0' ,BRAND
    union all
    Select 'MSO' AS FF, MSOTR as FFTR,BRAND, SUM(extinvmisc) AS SalesAmount   from V_FF_Brand_Sales 
    where TRANSDATE <= @LastDay
    group by MSOTR,BRAND
    ) as FFSales 
    ON (FFTarget.FFTR=FFSales.FFTR) AND (FFTarget.BRAND=FFSales.BRAND)
    group by FFTarget.FFTR
    ) as T1 
    order by FFTR asc """, conn.azure)
    sales_trend_data.to_excel("./Data/SalesTrend/sales_trend_data.xlsx", index=False)

    data = pd.read_excel('./Data/SalesTrend/sales_trend_data.xlsx')

    CBU = data[data['FFTR'].str.contains('CBU')]
    CBU.to_excel('./Data/SalesTrend/SalesTrend_CBU.xlsx', index=False)

    CCF = data[data['FFTR'].str.contains('CCF')]
    CCF.to_excel('./Data/SalesTrend/SalesTrend_CCF.xlsx', index=False)

    CCX = data[data['FFTR'].str.contains('CCX')]
    CCX.to_excel('./Data/SalesTrend/SalesTrend_CCX.xlsx', index=False)

    CNH = data[data['FFTR'].str.contains('CNH')]
    CNH.to_excel('./Data/SalesTrend/SalesTrend_CNH.xlsx', index=False)

    CKJ = data[data['FFTR'].str.contains('CKJ')]
    CKJ.to_excel('./Data/SalesTrend/SalesTrend_CKJ.xlsx', index=False)

    CMT = data[data['FFTR'].str.contains('CMT')]
    CMT.to_excel('./Data/SalesTrend/SalesTrend_CMT.xlsx', index=False)

    CRB = data[data['FFTR'].str.contains('CRB')]
    CRB.to_excel('./Data/SalesTrend/SalesTrend_CRB.xlsx', index=False)

    CRP = data[data['FFTR'].str.contains('CRP')]
    CRP.to_excel('./Data/SalesTrend_CRP.xlsx', index=False)

    CSB = data[data['FFTR'].str.contains('CSB')]
    CSB.to_excel('./Data/SalesTrend/SalesTrend_CSB.xlsx', index=False)

    print('3. Sales Trend Data Saved')
