import pandas as pd
import AdditonalFiles.db_connection as conn


def seen_rx_data():
    new_seen_rx_data = pd.read_sql_query(""" select * from
            (Select
            left([FF ID],3) as [FFTR],
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
            where [FF ID]  = left([FF ID],4)+'0' --and left([FF ID],3) like '%CBU%'
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
            from v_LastDay_SeenRx --where  left([FF ID],3) like '%CBU%'
                                        
            ) as T1
            order by [FFTR] asc """, conn.m_reporting)
    new_seen_rx_data.to_excel("./Data/SeenRx/Seen_Rx_Data.xlsx", index=False)

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
    CRP.to_excel('./Data/SeenRx/SeenRx_CRP.xlsx', index=False)

    CSB = data[data['FFTR'].str.contains('CSB')]
    CSB.to_excel('./Data/SeenRx/SeenRx_CSB.xlsx', index=False)
    print('2. Seen Rx Data Saved')


def doctor_call_data():
    doctor_call_data = pd.read_sql_query(""" 
        select * from
        (Select
        left([FF ID],3) as [FFTR],
        isnull(sum([LIGAZID Call]),0)  as [LIGAZID],
        isnull(sum([EMAZID Call]),0) as [EMAZID],
        isnull(sum([LIPICON Call]),0) as [LIPICON],
        isnull(sum([AGLIP Call]),0) as [AGLIP],
        isnull(sum([CIFIBET Call]),0) as [CIFIBET],
        isnull(sum([AMLEVO Call]),0) as [AMLEVO],
        isnull(sum([CARDOBIS Call]),0) as [CARDOBIS],
        isnull(sum([RIVAROX Call]),0) as [RIVAROX],
        isnull(sum([NOCLOG Call]),0) as [NOCLOG]
        from [V_LastDayDoctorCall] 
        where [FF ID]  = left([FF ID],4)+'0'
        group by left([FF ID],3)
                                                
        union all
                                                
        Select [FF ID],
        isnull([LIGAZID Call],0) as [LIGAZID Call],
        isnull([EMAZID Call],0) as [EMAZID Call],
        ISNULL([LIPICON Call], 0) as [LIPICON Call],
        ISNULL([AGLIP Call], 0) as [AGLIP Call],
        ISNULL([CIFIBET Call], 0) as [CIFIBET Call],
        ISNULL([AMLEVO Call], 0) as [AMLEVO Call],
        ISNULL([CARDOBIS Call], 0) as [CARDOBIS Call],
        ISNULL([RIVAROX Call], 0) as [RIVAROX Call],
        ISNULL([NOCLOG Call], 0) as [NOCLOG Call]
        from [V_LastDayDoctorCall] 
                                                
        ) as T1
        order by [FFTR] asc """, conn.m_reporting)
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
    CRP.to_excel('./Data/Call/Call_CRP.xlsx', index=False)

    CSB = data[data['FFTR'].str.contains('CSB')]
    CSB.to_excel('./Data/Call/Call_CSB.xlsx', index=False)

    print('3. Doctor Call Data Saved')


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
    CRP.to_excel('./Data/SalesTrend/SalesTrend_CRP.xlsx', index=False)

    CSB = data[data['FFTR'].str.contains('CSB')]
    CSB.to_excel('./Data/SalesTrend/SalesTrend_CSB.xlsx', index=False)

    print('1. Sales Trend Data Saved')
