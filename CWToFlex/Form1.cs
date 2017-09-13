using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading;
using System.Globalization;
using System.IO;

namespace CWToFlex
{
    public partial class Form1 : Form
    {
        static DateTime dFirstDayOfThisMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));
        DateTime dLastDayOfLastMonth = dFirstDayOfThisMonth.AddDays(-1);
        DateTime dFirstDayOfLastMonth = dFirstDayOfThisMonth.AddMonths(-1);
        public Form1()
        {
            // Sets the culture to Swedish
            Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-SE");
            // Sets the UI culture to Swedish
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("sv-SE");
            
            InitializeComponent();
            btnCreateFile.Visible = false;
            DateTime dFirstDayOfThisMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));
            DateTime dLastDayOfLastMonth = dFirstDayOfThisMonth.AddDays(-1);
            DateTime dFirstDayOfLastMonth = dFirstDayOfThisMonth.AddMonths(-1);
            dtpFrom.Value = dFirstDayOfLastMonth;
            dtpTom.Value = dLastDayOfLastMonth;
            txtSaveFileToDir.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Flex";
            DirectoryInfo di = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Flex");
            if (!di.Exists)
                di.Create();
        }

        private void btnGetCWData_Click(object sender, EventArgs e)
        {
            ss1.Items.Clear();
            ss1.ForeColor = Color.Red;
            ss1.Items.Add("Hämtar data från CW SQL databas, vänligen vänta........................................");
            lbCWData.Items.Clear();
            lbCWData.Items.Add("Version: 1.3 Acon CW data till Flex lön " + DateTime.Now.ToString());
            lbCWData.Items.Add(dtpFrom.Value.AddDays(24).ToShortDateString().Replace("-","") + "\t" +
               dtpFrom.Value.ToShortDateString().Replace("-","") + "\t" +
                dtpTom.Value.ToShortDateString().Replace("-", ""));
            //Deviation and Overtime
            foreach (DataRow row in loadDataTable(getSelectStringDeviationOvertime(dtpFrom.Value, dtpTom.Value), "Deviation").Rows)
            {
                if (row["EmployeeNo"].ToString() != "0" || row["EmployeeNo"].ToString() == "")
                {
                    lbCWData.Items.Add(row["EmployeeNo"] //Anstnr
                      + "\t"
                      + row["Loneartsnr"] //Löneartsnr
                      + "\t"
                      + row["Account"] //Konteringsnivå 1
                      + "\t"
                      + "" //Konteringsnivå 2
                      + "\t"
                      + "" //Konteringsnivå 3
                      + "\t"
                      + "" //Konteringsnivå 4
                      + "\t"
                      + "" //Konteringsnivå 5
                      + "\t"
                      + "" //Konteringsnivå 6
                      + "\t"
                      + "" //Konteringsnivå 7
                      + "\t"
                      + "" //Konteringsnivå 8
                      + "\t"
                      + "" //Konteringsnivå 9
                      + "\t"
                      + "" //Konteringsnivå 10
                      + "\t"
                      + row["Hours"] //Antal
                      + "\t"
                      + "tim" //Antal enhet Enheten för antal, kan vara 'tim', 'dgr', ’kdgr’, eller utelämnas  
                      + "\t"
                      + "" //A-pris
                      + "\t"
                      + "" //Belopp
                      + "\t"
                      // + "" //
                      //+ "\t"
                      + Convert.ToDateTime(row["Date_Start"]).ToShortDateString().Replace("-", "")  //Fr.o.m.datum
                      + "\t"
                      + Convert.ToDateTime(row["Date_End"]).ToShortDateString().Replace("-", "") //T.o.m. datum
                      + "\t"
                      +   row["ActivityType"] //Meddelande
                      + "\t"
                      + "" //Omfattning %
                      + "\t"
                      + "" //Lönekod
                      + "\t"
                      + "" //Semesterkvot
                      + "\t"
                      + "" //Kalenderdagsfaktor
                       + "\t"
                      + "" //Barn //
                       + "\t"
                      );
                }
                else
                    
                    lbCWData.Items.Add(row["MemberName"] //Anstnr
                    + "\t"
                    + row["Loneartsnr"] //Löneartsnr
                    + "\t"
                    + row["Account"] //Konteringsnivå 1
                    + "\t"
                    + "" //Konteringsnivå 2
                    + "\t"
                    + "" //Konteringsnivå 3
                    + "\t"
                    + "" //Konteringsnivå 4
                    + "\t"
                    + "" //Konteringsnivå 5
                    + "\t"
                    + "" //Konteringsnivå 6
                    + "\t"
                    + "" //Konteringsnivå 7
                    + "\t"
                    + "" //Konteringsnivå 8
                    + "\t"
                    + "" //Konteringsnivå 9
                    + "\t"
                    + "" //Konteringsnivå 10
                    + "\t"
                    + row["Hours"] //Antal
                    + "\t"
                    + "tim" //Antal enhet Enheten för antal, kan vara 'tim', 'dgr', ’kdgr’, eller utelämnas  
                    + "\t"
                    + "" //A-pris
                    + "\t"
                    + "" //Belopp
                    //+ "\t"
                    //+ "" //
                    + "\t"
                    + Convert.ToDateTime(row["Date_Start"]).ToShortDateString().Replace("-", "")  //Fr.o.m.datum
                    + "\t"
                    + Convert.ToDateTime(row["Date_End"]).ToShortDateString().Replace("-", "") //T.o.m. datum
                    + "\t"
                    +  row["ActivityType"] //Meddelande
                    + "\t"
                    + "" //Omfattning %
                    + "\t"
                    + "" //Lönekod
                    + "\t"
                    + "" //Semesterkvot
                    + "\t"
                    + "" //Kalenderdagsfaktor
                     + "\t"
                    + "" //Barn
                     + "\t"
                    );
            }
                //Expenses
            foreach (DataRow row in loadDataTable(getSelectStringExpense(dtpFrom.Value, dtpTom.Value), "Expense").Rows)
            {
                if (row["EmployeeNo"].ToString() != "0" || row["EmployeeNo"].ToString() == "")
                    lbCWData.Items.Add(row["EmployeeNo"] //Anstnr
                    + "\t"
                    + row["Loneartsnr"] //Löneartsnr
                    + "\t"
                    + row["Account"] //Konteringsnivå 1
                    + "\t"
                    + "" //Konteringsnivå 2
                    + "\t"
                    + "" //Konteringsnivå 3
                    + "\t"
                    + "" //Konteringsnivå 4
                    + "\t"
                    + "" //Konteringsnivå 5
                    + "\t"
                    + "" //Konteringsnivå 6
                    + "\t"
                    + "" //Konteringsnivå 7
                    + "\t"
                    + "" //Konteringsnivå 8
                    + "\t"
                    + "" //Konteringsnivå 9
                    + "\t"
                    + "" //Konteringsnivå 10
                    + "\t"
                    + row["Quantity"] //Antal
                    + "\t"
                    + "" //Antal enhet Enheten för antal, kan vara 'tim', 'dgr', ’kdgr’, eller utelämnas  
                    + "\t"
                    + "" //A-pris
                    + "\t"
                    + row["Amount"] //Belopp
                    + "\t"
                    + "" //
                    //+ "\t"
                    + Convert.ToDateTime(row["Date_Expense"]).ToShortDateString().Replace("-", "")  //Fr.o.m.datum
                    + "\t"
                    + Convert.ToDateTime(row["Date_Expense"]).ToShortDateString().Replace("-", "") //T.o.m. datum
                    + "\t"
                    + row["ETDescription"] //Meddelande
                    + "\t"
                    + "" //Omfattning %
                    + "\t"
                    + "" //Lönekod
                    + "\t"
                    + "" //Semesterkvot
                    + "\t"
                    + "" //Kalenderdagsfaktor
                    + "\t"
                    + "" //Barn
                    );
                else
                    lbCWData.Items.Add(row["MemberName"] //Anstnr
                    + "\t"
                    + row["Loneartsnr"] //Löneartsnr
                    + "\t"
                    + row["Account"] //Konteringsnivå 1
                    + "\t"
                    + "" //Konteringsnivå 2
                    + "\t"
                    + "" //Konteringsnivå 3
                    + "\t"
                    + "" //Konteringsnivå 4
                    + "\t"
                    + "" //Konteringsnivå 5
                    + "\t"
                    + "" //Konteringsnivå 6
                    + "\t"
                    + "" //Konteringsnivå 7
                    + "\t"
                    + "" //Konteringsnivå 8
                    + "\t"
                    + "" //Konteringsnivå 9
                    + "\t"
                    + "" //Konteringsnivå 10
                    + "\t"
                    + row["Quantity"] //Antal
                    + "\t"
                    + "" //Antal enhet Enheten för antal, kan vara 'tim', 'dgr', ’kdgr’, eller utelämnas  
                    + "\t"
                    + "" //A-pris
                    + "\t"
                    + row["Amount"] //Belopp
                    + "\t"
                    + "" //
                   // + "\t"
                    + Convert.ToDateTime(row["Date_Expense"]).ToShortDateString().Replace("-", "")  //Fr.o.m.datum
                    + "\t"
                    + Convert.ToDateTime(row["Date_Expense"]).ToShortDateString().Replace("-", "") //T.o.m. datum
                    + "\t"
                    +  row["ETDescription"]  //Meddelande
                    + "\t"
                    + "" //Omfattning %
                    + "\t"
                    + "" //Lönekod
                    + "\t"
                    + "" //Semesterkvot
                    + "\t"
                    + "" //Kalenderdagsfaktor
                    + "\t"
                    + "" //Barn
                    );
                btnCreateFile.Visible = true;
            }
            ss1.Items.Clear();
            ss1.ForeColor = Color.Green;
            ss1.Items.Add("Hämtning av data från CW SQL databas har slutförts");
        }
        private DataTable loadDataTable(string Select, string TableName)
        {
            string connString ="Database=cwwebapp_acon;Server=Aconsql02;User=SSRSAcon;Password=SSRSAcon;connect timeout = 30";
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(Select, conn);
            conn.Open();

            // create data adapter
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable
            DataTable dt = new DataTable(TableName);
            da.Fill(dt);
            conn.Close();
            da.Dispose();
            return dt;
        }
        private string getSelectStringExpense(DateTime From, DateTime Tom)
        {
            //Killemeterersättning dela upp per konto
            //Representation konto 6071 och 7631
            string select = @"SELECT m.EmployeeNo           
                            ,ed.Reason, ed.Date_Expense           
                            ,case WHEN et.EX_Type_RecID = 35 and et.Description like 'Kilometerersättning'  THEN Amount            
                            when ed.Bill_Amount = 0 and et.Description like 'Kilometerersättning' then ed.Invoice_Amount           
                            when ed.Invoice_Amount = 0 and et.Description like 'Kilometerersättning' then ed.Bill_Amount           
                            else 0 end Quantity           , case WHEN et.EX_Type_RecID = 35 and et.Description not like 'Kilometerersättning' THEN Amount            
                            when ed.Bill_Amount = 0 and et.Description not like 'Kilometerersättning' then ed.Invoice_Amount           
                            when ed.Invoice_Amount = 0 and et.Description not like 'Kilometerersättning' then ed.Bill_Amount           
                            when  et.Description not like 'Kilometerersättning' then ed.Invoice_Amount else 0 end Amount           
                            ,CASE WHEN et.EX_Type_RecID = 35 THEN et.Description + ' (' + CAST(CAST(ed.Amount AS int) AS varchar) + ' km)' ELSE et.Description END AS ETDescription           
                            , CASE WHEN Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0 THEN 'Acon'           
                            WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 THEN 'Utlägg'           
                            END AS Accounting           ,et.Integration_Xref Loneartsnr           
                            ,CASE WHEN (Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0 ) AND et.Description not like 'Traktament%' THEN 'xxxx'           
                            --WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Flygbiljett' THEN '5810'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Parkering' THEN '5822'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Taxi/Tåg/Båt/Buss' THEN '5811'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Hyrbil' THEN '5820'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Delar och tillbehör' THEN '5460'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Övrigt' THEN '5460'           
                            --WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Drivmedel' THEN '5611'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Reparation' THEN '5520'           
                            --WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Backuplicens' THEN '4080'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Programvaror' THEN '5420'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Utbildning' THEN '7610'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Tull' THEN 'x'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Kilometerersättning' THEN '7331'           
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Företags Representation' THEN '6071'          
                            -- WHEN(Company_Name LIKE 'Acon AB' OR ed.Billable_Flag = 0) AND et.Description LIKE 'Personal Representation' THEN '7631'           
                            WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description not like 'Traktament%' THEN 'yyyy'          
                            WHEN et.Description like '%Sverige' THEN '7321'           
                            WHEN et.Description like '%SKP' THEN '7322'           
                            WHEN et.Description like '%UTLAND' THEN '7323'           
                            WHEN et.Description like 'Traktament%' THEN ''          
                            -- WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Flygbiljett' THEN '4029'          
                            -- WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Parkering' THEN '4029'          
                            -- WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Taxi/Tåg/Båt/Buss' THEN '4029'          
                            -- WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Hyrbil' THEN '4029'          
                            -- WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Delar och tillbehör' THEN '4029'          
                            -- WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Övrigt' THEN '4029'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Drivmedel' THEN '4029'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Reparation' THEN '4029'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Backuplicens' THEN '4029'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Programvaror' THEN '4029'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Utbildning' THEN '4029'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Tull' THEN '4029'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Kilometerersättning' THEN '7331'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Företags Representation' THEN '6071'         
                            --  WHEN Company_Name NOT LIKE 'Acon AB' AND ed.Billable_Flag = 1 AND et.Description LIKE 'Personal Representation' THEN '7631'            
                            ELSE 'Fel'           
                            END AS Account           
                            FROM EX_Header AS e INNER JOIN           
                            EX_Detail AS ed ON ed.EX_Header_RecID = e.EX_Header_RecID INNER JOIN           
                            EX_Class AS ec ON ec.EX_Class_RecID = ed.EX_Class_RecID INNER JOIN           
                            EX_Payment AS ep ON ep.EX_Payment_RecID = ed.EX_Payment_RecID INNER JOIN           
                            (Select m.*, me.Value as EmployeeNo from Member m           
                            left join Member_Extended_Property_Value me on me.Member_RecID = m.Member_RecID and me.Member_Extended_Property_RecID = 98) m ON m.Member_RecID = e.Member_RecID INNER JOIN           
                            Company AS c ON c.Company_RecID = ed.Company_RecID INNER JOIN           
                            EX_Type AS et ON et.EX_Type_RecID = ed.EX_Type_RecID           
                            WHERE(ed.Date_Expense BETWEEN '" + From.ToShortDateString() + "' and '" + Tom.ToShortDateString() + @"')
                            and ec.EX_Class_RecID = 2
                            order by cast(m.EmployeeNo as int)";
            return select;
        }
        private string getSelectStringDeviationOvertime(DateTime From, DateTime Tom)
        {
            string select = @" -- 36\tInternt: VAB                            
-- 42\tInternt: Semester                            
-- 43\tInternt: Föräldraledig                            
-- 45\tInternt: Pappadagar                            
-- 46\tInternt: Tjänstledig                            
Select                             
m.Member_RecID as MemberKey                            
,cast(m.EmployeeNo as int) EmployeeNo                            
,m.First_Name + ' ' + m.Last_Name as MemberName                            
,case when Date_start_1 is null then Date_End_1                            
when Date_start_2 is null then Date_Start_1                            
when Date_start_3 is null then Date_Start_2                            
when Date_start_4 is null then Date_Start_3                            
when Date_start_5 is null then Date_Start_4                            
when Date_start_6 is null then Date_Start_5                            
when Date_start_7 is null then Date_Start_6                            
end Date_Start                            
,case when Date_End_2 is null then Date_End_1                            
when Date_End_3 is null then Date_End_2                            
when Date_End_4 is null then Date_End_3                            
when Date_End_5 is null then Date_End_4                            
when Date_End_6 is null then Date_End_5                            
when Date_End_7 is null then Date_End_6                            
end Date_End                            
, at.Description as ActivityType                            
,0 as IsOverTime                            
,0 as OverTime                            
--,( case when datepart(WEEKDAY,t.Date_Start) = 1 or datepart(WEEKDAY,t.Date_Start) = 7 then 0 else 1 end) as IsWorkDay                            
,case when h.Holiday_Date is not null then 1 else 0 end as IsHoliday                            
,sum(t.Hours_Actual) Hours                            
,Case when sum(t.Hours_Actual) < 8 and at.Description like 'Internt: Föräldraledig' then 613 else  at.Integration_Xref end LoneArtsNr                            
,at.Xref_Work_Type Account                            
from Time_Entry t                             
inner join Time_Sheet ts on ts.Time_Sheet_RecID = t.Time_Sheet_RecID                            
inner join TE_Period tp on tp.TE_Period_RecID = ts.TE_Period_RecID                            
inner join (Select m.*,me.Value as EmployeeNo from Member m                             
left join Member_Extended_Property_Value me on me.Member_RecID = m.Member_RecID and me.Member_Extended_Property_RecID = 98) m on m.Member_RecID = t.Member_RecID                            
inner join Activity_Type at on at.Activity_Type_RecID = t.Activity_Type_RecID                            left outer join holiday h on convert(varchar,h.Holiday_Date,112) = convert(varchar,t.Date_Start,112)
--StartDatum                            
outer apply(select max(date_start)Date_start_1 from Time_Entry
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,t.date_Start) <= 2 and Member_RecID = t.Member_RecID and date_start < t.Date_Start and Activity_Type_RecID = t.Activity_Type_RecID) ts1                            
outer apply(select max(date_start)Date_start_2 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts1.Date_start_1) <= 2 and Member_RecID = t.Member_RecID and date_start < ts1.Date_Start_1 and Activity_Type_RecID = t.Activity_Type_RecID) ts2                            
outer apply(select max(date_start)Date_start_3 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts2.Date_start_2) <= 2  and Member_RecID = t.Member_RecID and date_start < ts2.Date_Start_2 and Activity_Type_RecID = t.Activity_Type_RecID) ts3                            
outer apply(select max(date_start)Date_start_4 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts3.Date_start_3) <= 2  and Member_RecID = t.Member_RecID and date_start < ts3.Date_Start_3 and Activity_Type_RecID = t.Activity_Type_RecID) ts4                            
outer apply(select max(date_start)Date_start_5 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts4.Date_start_4) <= 2  and Member_RecID = t.Member_RecID and date_start < ts4.Date_Start_4 and Activity_Type_RecID = t.Activity_Type_RecID) ts5                            
outer apply(select max(date_start)Date_start_6 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts5.Date_start_5) <= 2   and Member_RecID = t.Member_RecID and date_start < ts5.Date_Start_5 and Activity_Type_RecID = t.Activity_Type_RecID) ts6                            
outer apply(select max(date_start)Date_start_7 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts6.Date_start_6) <= 2   and Member_RecID = t.Member_RecID and date_start < ts6.Date_Start_6 and Activity_Type_RecID = t.Activity_Type_RecID) ts7                            
outer apply(select max(date_start)Date_start_8 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts7.Date_start_7) <= 2   and Member_RecID = t.Member_RecID and date_start < ts7.Date_Start_7 and Activity_Type_RecID = t.Activity_Type_RecID) ts8                            
outer apply(select max(date_start)Date_start_9 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts8.Date_start_8) <= 2   and Member_RecID = t.Member_RecID and date_start < ts8.Date_Start_8 and Activity_Type_RecID = t.Activity_Type_RecID) ts9                            
outer apply(select max(date_start)Date_start_10 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts9.Date_start_9) <= 2   and Member_RecID = t.Member_RecID and date_start < ts9.Date_Start_9 and Activity_Type_RecID = t.Activity_Type_RecID) ts10                            
outer apply(select max(date_start)Date_start_11 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts10.Date_start_10) <= 2   and Member_RecID = t.Member_RecID and date_start < ts10.Date_Start_10 and Activity_Type_RecID = t.Activity_Type_RecID) ts11                            
outer apply(select max(date_start)Date_start_12 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts11.Date_start_11) <= 2   and Member_RecID = t.Member_RecID and date_start < ts11.Date_Start_11 and Activity_Type_RecID = t.Activity_Type_RecID) ts12                            
outer apply(select max(date_start)Date_start_13 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts12.Date_start_12) <= 2   and Member_RecID = t.Member_RecID and date_start < ts12.Date_Start_12 and Activity_Type_RecID = t.Activity_Type_RecID) ts13                            
outer apply(select max(date_start)Date_start_14 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts13.Date_start_13) <= 2   and Member_RecID = t.Member_RecID and date_start < ts13.Date_Start_13 and Activity_Type_RecID = t.Activity_Type_RecID) ts14                            
outer apply(select max(date_start)Date_start_15 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts14.Date_start_14) <= 2   and Member_RecID = t.Member_RecID and date_start < ts14.Date_Start_14 and Activity_Type_RecID = t.Activity_Type_RecID) ts15                            
outer apply(select max(date_start)Date_start_16 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts15.Date_start_15) <= 2   and Member_RecID = t.Member_RecID and date_start < ts15.Date_Start_15 and Activity_Type_RecID = t.Activity_Type_RecID) ts16                            
outer apply(select max(date_start)Date_start_17 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts16.Date_start_16) <= 2   and Member_RecID = t.Member_RecID and date_start < ts16.Date_Start_16 and Activity_Type_RecID = t.Activity_Type_RecID) ts17                            
outer apply(select max(date_start)Date_start_18 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts17.Date_start_17) <= 2   and Member_RecID = t.Member_RecID and date_start < ts17.Date_Start_17 and Activity_Type_RecID = t.Activity_Type_RecID) ts18                            
outer apply(select max(date_start)Date_start_19 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts18.Date_start_18) <= 2   and Member_RecID = t.Member_RecID and date_start < ts18.Date_Start_18 and Activity_Type_RecID = t.Activity_Type_RecID) ts19                            
outer apply(select max(date_start)Date_start_20 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts19.Date_start_19) <= 2   and Member_RecID = t.Member_RecID and date_start < ts19.Date_Start_19 and Activity_Type_RecID = t.Activity_Type_RecID) ts20                            
outer apply(select max(date_start)Date_start_21 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts20.Date_start_20) <= 2   and Member_RecID = t.Member_RecID and date_start < ts20.Date_Start_20 and Activity_Type_RecID = t.Activity_Type_RecID) ts21                            
outer apply(select max(date_start)Date_start_22 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts21.Date_start_21) <= 2   and Member_RecID = t.Member_RecID and date_start < ts21.Date_Start_21 and Activity_Type_RecID = t.Activity_Type_RecID) ts22                            
outer apply(select max(date_start)Date_start_23 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts22.Date_start_22) <= 2   and Member_RecID = t.Member_RecID and date_start < ts22.Date_Start_22 and Activity_Type_RecID = t.Activity_Type_RecID) ts23                            
outer apply(select max(date_start)Date_start_24 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts23.Date_start_23) <= 2   and Member_RecID = t.Member_RecID and date_start < ts23.Date_Start_23 and Activity_Type_RecID = t.Activity_Type_RecID) ts24                            
outer apply(select max(date_start)Date_start_25 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts24.Date_start_24) <= 2   and Member_RecID = t.Member_RecID and date_start < ts24.Date_Start_24 and Activity_Type_RecID = t.Activity_Type_RecID) ts25                            
outer apply(select max(date_start)Date_start_26 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts25.Date_start_25) <= 2   and Member_RecID = t.Member_RecID and date_start < ts25.Date_Start_25 and Activity_Type_RecID = t.Activity_Type_RecID) ts26                            
outer apply(select max(date_start)Date_start_27 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts26.Date_start_26) <= 2   and Member_RecID = t.Member_RecID and date_start < ts26.Date_Start_26 and Activity_Type_RecID = t.Activity_Type_RecID) ts27                            
outer apply(select max(date_start)Date_start_28 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts27.Date_start_27) <= 2   and Member_RecID = t.Member_RecID and date_start < ts27.Date_Start_27 and Activity_Type_RecID = t.Activity_Type_RecID) ts28                            
outer apply(select max(date_start)Date_start_29 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts28.Date_start_28) <= 2   and Member_RecID = t.Member_RecID and date_start < ts28.Date_Start_28 and Activity_Type_RecID = t.Activity_Type_RecID) ts29                            
outer apply(select max(date_start)Date_start_30 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datepart(Year,date_start) = datepart(year,t.date_start) and datediff(day,date_start,ts29.Date_start_29) <= 2   and Member_RecID = t.Member_RecID and date_start < ts29.Date_Start_29 and Activity_Type_RecID = t.Activity_Type_RecID) ts30                            
outer apply(select max(date_start)Date_start_31 from Time_Entry                            
where datepart(month,date_start) = datepart(month,t.date_start) and datediff(day,date_start,ts30.Date_start_30) <= 2   and Member_RecID = t.Member_RecID and date_start < ts30.Date_Start_30 and Activity_Type_RecID = t.Activity_Type_RecID) ts31                            
-- Slutdatum                            
outer apply(select min(date_Start)Date_End_1 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and datediff(day,date_start,t.Date_start) <=2   and Member_RecID = t.Member_RecID and date_Start >= t.Date_Start and Activity_Type_RecID = t.Activity_Type_RecID) t1                            
outer apply(select min(date_Start)Date_End_2 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and datediff(day,t1.Date_End_1,Date_Start) <= 2   and Member_RecID = t.Member_RecID and date_Start > t1.Date_End_1 and Activity_Type_RecID = t.Activity_Type_RecID) t2                            
outer apply(select min(date_Start)Date_End_3 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start)  and datediff(day,t2.Date_End_2,Date_Start) <= 2   and Member_RecID = t.Member_RecID and date_Start > t2.Date_End_2 and Activity_Type_RecID = t.Activity_Type_RecID) t3                            
outer apply(select min(date_Start)Date_End_4 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start)  and datediff(day,t3.Date_End_3,Date_Start) <= 2  and Member_RecID = t.Member_RecID and date_Start > t3.Date_End_3 and Activity_Type_RecID = t.Activity_Type_RecID) t4                            
outer apply(select min(date_Start)Date_End_5 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and datediff(day,t4.Date_End_4,Date_Start) <= 2  and Member_RecID = t.Member_RecID and date_Start > t4.Date_End_4 and Activity_Type_RecID = t.Activity_Type_RecID) t5                            
outer apply(select min(date_Start)Date_End_6 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and datediff(day,t5.Date_End_5,Date_Start) <= 2 and Member_RecID = t.Member_RecID and date_Start > t5.Date_End_5 and Activity_Type_RecID = t.Activity_Type_RecID) t6                            
outer apply(select min(date_Start)Date_End_7 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and datediff(day,t6.Date_End_6,Date_Start) <= 2 and Member_RecID = t.Member_RecID and date_Start > t6.Date_End_6 and Activity_Type_RecID = t.Activity_Type_RecID) t7                            
outer apply(select min(date_Start)Date_End_8 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and datediff(day,t7.Date_End_7,Date_Start) <= 2 and Member_RecID = t.Member_RecID and date_Start > t7.Date_End_7 and Activity_Type_RecID = t.Activity_Type_RecID) t8                            
outer apply(select min(date_Start)Date_End_9 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and datediff(day,t8.Date_End_8,Date_Start) <= 2 and Member_RecID = t.Member_RecID and date_Start > t8.Date_End_8 and Activity_Type_RecID = t.Activity_Type_RecID) t9                            
outer apply(select min(date_Start)Date_End_10 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and datediff(day,t9.Date_End_9,Date_Start) <= 2 and Member_RecID = t.Member_RecID and date_Start > t9.Date_End_9 and Activity_Type_RecID = t.Activity_Type_RecID) t10                            
outer apply(select min(date_Start)Date_End_11 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t10.Date_End_10 and Activity_Type_RecID = t.Activity_Type_RecID) t11                            
outer apply(select min(date_Start)Date_End_12 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t11.Date_End_11 and Activity_Type_RecID = t.Activity_Type_RecID) t12                            
outer apply(select min(date_Start)Date_End_13 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t12.Date_End_12 and Activity_Type_RecID = t.Activity_Type_RecID) t13                            
outer apply(select min(date_Start)Date_End_14 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t13.Date_End_13 and Activity_Type_RecID = t.Activity_Type_RecID) t14                            
outer apply(select min(date_Start)Date_End_15 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t14.Date_End_14 and Activity_Type_RecID = t.Activity_Type_RecID) t15                            
outer apply(select min(date_Start)Date_End_16 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t15.Date_End_15 and Activity_Type_RecID = t.Activity_Type_RecID) t16                            
outer apply(select min(date_Start)Date_End_17 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t16.Date_End_16 and Activity_Type_RecID = t.Activity_Type_RecID) t17                            
outer apply(select min(date_Start)Date_End_18 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t17.Date_End_17 and Activity_Type_RecID = t.Activity_Type_RecID) t18                            
outer apply(select min(date_Start)Date_End_19 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t18.Date_End_18 and Activity_Type_RecID = t.Activity_Type_RecID) t19                            
outer apply(select min(date_Start)Date_End_20 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t19.Date_End_19 and Activity_Type_RecID = t.Activity_Type_RecID) t20                            
outer apply(select min(date_Start)Date_End_21 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t20.Date_End_20 and Activity_Type_RecID = t.Activity_Type_RecID) t21                            
outer apply(select min(date_Start)Date_End_22 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t21.Date_End_21 and Activity_Type_RecID = t.Activity_Type_RecID) t22                            
outer apply(select min(date_Start)Date_End_23 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t22.Date_End_22 and Activity_Type_RecID = t.Activity_Type_RecID) t23                            
outer apply(select min(date_Start)Date_End_24 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t23.Date_End_23 and Activity_Type_RecID = t.Activity_Type_RecID) t24                            
outer apply(select min(date_Start)Date_End_25 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t24.Date_End_24 and Activity_Type_RecID = t.Activity_Type_RecID) t25                            
outer apply(select min(date_Start)Date_End_26 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t25.Date_End_25 and Activity_Type_RecID = t.Activity_Type_RecID) t26                            
outer apply(select min(date_Start)Date_End_27 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t26.Date_End_26 and Activity_Type_RecID = t.Activity_Type_RecID) t27                            
outer apply(select min(date_Start)Date_End_28 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t27.Date_End_27 and Activity_Type_RecID = t.Activity_Type_RecID) t28                            
outer apply(select min(date_Start)Date_End_29 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t28.Date_End_28 and Activity_Type_RecID = t.Activity_Type_RecID) t29                            
outer apply(select min(date_Start)Date_End_30 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t29.Date_End_29 and Activity_Type_RecID = t.Activity_Type_RecID) t30                            
outer apply(select min(date_Start)Date_End_31 from Time_Entry                            
where datepart(month,date_Start) = datepart(month,t.date_Start) and datepart(Year,date_Start) = datepart(year,t.date_Start) and Member_RecID = t.Member_RecID and date_Start > t30.Date_End_30 and Activity_Type_RecID = t.Activity_Type_RecID) t31                            
--left join Time_Entry t1 on t1.Member_RecID = t.Member_RecID and t1.Activity_Type_RecID = 6 and datepart(month,t1.date_start) = datepart(month,t.date_start) and datediff(day,t.Date_Start, t1.Date_Start) between 1 and 4                            
where ((t.Date_Start between '" + From.ToShortDateString() + "' and '" + Tom.ToShortDateString() + @"' --'" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'                             
and at.Utilization_Flag = 0 and at.Inactive_Flag = 0 and at.Activity_Type_RecID not in (6,7) and m.Member_Type_RecID not in (6) ))-- and t.Member_ID like 'DArlebrandt'  ) --NHiltunen JAndersson ABjorn LEriksson DArlebrandt                            
group by m.Member_RecID                            
,cast(m.EmployeeNo as int)                             
,m.First_Name + ' ' + m.Last_Name                             
,case when Date_start_1 is null then Date_End_1                            
when Date_start_2 is null then Date_Start_1                            
when Date_start_3 is null then Date_Start_2                            
when Date_start_4 is null then Date_Start_3                            
when Date_start_5 is null then Date_Start_4                            
when Date_start_6 is null then Date_Start_5                            
when Date_start_7 is null then Date_Start_6                            
end                             
,case when Date_End_2 is null then Date_End_1                            
when Date_End_3 is null then Date_End_2                            
when Date_End_4 is null then Date_End_3                            
when Date_End_5 is null then Date_End_4                            
when Date_End_6 is null then Date_End_5                            
when Date_End_7 is null then Date_End_6                            
end                            
, at.Description                             
,case when h.Holiday_Date is not null then 1 else 0 end                             
,at.Integration_Xref                             
,at.Xref_Work_Type                             
union all                             
-- 7 Kompledighet                            
Select                             
m.Member_RecID as MemberKey                            
,cast(m.EmployeeNo as int) EmployeeNo                            
,m.First_Name + ' ' + m.Last_Name as MemberName                            
,t.Date_Start                            
,t.Date_Start Date_End                            
, at.Description as ActivityType                            
,0 as IsOverTime                            
,0 as OverTime                            
--,( case when datepart(WEEKDAY,t.Date_Start) = 1 or datepart(WEEKDAY,t.Date_Start) = 7 then 0 else 1 end) as IsWorkDay                            
,case when h.Holiday_Date is not null then 1 else 0 end as IsHoliday                            
,sum(t.Hours_Actual) Hours                            
,Case when sum(t.Hours_Actual) < 8 and at.Description like 'Internt: Föräldraledig' then 613 else  at.Integration_Xref end LoneArtsNr                            
,at.Xref_Work_Type Account                            
from Time_Entry t                            
inner join Time_Sheet ts on ts.Time_Sheet_RecID = t.Time_Sheet_RecID                            
inner join TE_Period tp on tp.TE_Period_RecID = ts.TE_Period_RecID                            
inner join (Select m.*,me.Value as EmployeeNo from Member m                             
left join Member_Extended_Property_Value me on me.Member_RecID = m.Member_RecID and me.Member_Extended_Property_RecID = 98) m on m.Member_RecID = t.Member_RecID                            
inner join Activity_Type at on at.Activity_Type_RecID = t.Activity_Type_RecID                            
left outer join holiday h on convert(varchar,h.Holiday_Date,112) = convert(varchar,t.Date_Start,112)                            
where ((t.Date_Start between '" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'                            
and at.Utilization_Flag = 0 and at.Inactive_Flag = 0 and at.Activity_Type_RecID in (7) )) and m.Member_Type_RecID not in (6)-- and t.Member_ID like 'DArlebrandt'  ) --NHiltunen JAndersson ABjorn LEriksson DArlebrandt
group by m.Member_RecID                            
,cast(m.EmployeeNo as int)                             
,m.First_Name + ' ' + m.Last_Name                            
,t.Date_Start                            
, at.Description                             
,case when h.Holiday_Date is not null then 1 else 0 end                             
,at.Integration_Xref                             
,at.Xref_Work_Type                            
union all                            
-- 6 sjukdom                            
Select                             
m.Member_RecID as MemberKey                            
,cast(m.EmployeeNo as int) EmployeeNo                            
,m.First_Name + ' ' + m.Last_Name as MemberName                            
,t.Date_Start                            
,t.Date_Start Date_End                            
, at.Description as ActivityType                            
,0 as IsOverTime                            
,0 as OverTime                            
--,( case when datepart(WEEKDAY,t.Date_Start) = 1 or datepart(WEEKDAY,t.Date_Start) = 7 then 0 else 1 end) as IsWorkDay                            
,case when h.Holiday_Date is not null then 1 else 0 end as IsHoliday                            
,sum(t.Hours_Actual) Hours                            
,Case when sum(t.Hours_Actual) < 8 and at.Description like 'Internt: Föräldraledig' then 613 else  at.Integration_Xref end LoneArtsNr                            
,at.Xref_Work_Type Account                            
from Time_Entry t                            
inner join Time_Sheet ts on ts.Time_Sheet_RecID = t.Time_Sheet_RecID                            
inner join TE_Period tp on tp.TE_Period_RecID = ts.TE_Period_RecID                            
inner join (Select m.*,me.Value as EmployeeNo from Member m                             
left join Member_Extended_Property_Value me on me.Member_RecID = m.Member_RecID and me.Member_Extended_Property_RecID = 98) m on m.Member_RecID = t.Member_RecID                            
inner join Activity_Type at on at.Activity_Type_RecID = t.Activity_Type_RecID                            
left outer join holiday h on convert(varchar,h.Holiday_Date,112) = convert(varchar,t.Date_Start,112)                            
where ((t.Date_Start between '" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'                            
and at.Utilization_Flag = 0 and at.Inactive_Flag = 0 and at.Activity_Type_RecID in (6) )) and m.Member_Type_RecID not in (6)-- and t.Member_ID like 'DArlebrandt'  ) --NHiltunen JAndersson ABjorn LEriksson DArlebrandt
group by m.Member_RecID                            
,cast(m.EmployeeNo as int)                             
,m.First_Name + ' ' + m.Last_Name                             
,t.Date_Start                            
, at.Description                             
,case when h.Holiday_Date is not null then 1 else 0 end                             
,at.Integration_Xref                             
,at.Xref_Work_Type                             
union all                              
--Timlön                            
Select                            
m.Member_RecID as MemberKey                            
,cast(m.EmployeeNo as int) EmployeeNo                            
,m.First_Name + ' ' + m.Last_Name as MemberName                
,'" + From.ToShortDateString() + @"'                
,'" + Tom.ToShortDateString() + @"'                            
,'Hourly Salery' as ActivityType                            
,0 as IsOverTime                            
,0 as OverTime                            
,case when h.Holiday_Date is not null then 1 else 0 end as IsHoliday                            
,sum(case                             
when h.Holiday_Date is not null then 0                            
when datepart(WEEKDAY, t.Date_Start) = 1 or datepart(WEEKDAY, t.Date_Start) = 7 then  0                            
when t.Time_Start >= '18:00:00' then 0                            
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' then 0                            
when t.Time_End >= '18:00:00' and t.Time_Start <= '18:00:00' then 0                            
else t.Hours_Actual end) as Hours                           
,at.Integration_Xref LoneArtsNr                            
, at.Xref_Work_Type Account                            
from Time_Entry t                            
inner join Time_Sheet ts on ts.Time_Sheet_RecID = t.Time_Sheet_RecID                            
inner join TE_Period tp on tp.TE_Period_RecID = ts.TE_Period_RecID                            
inner join (Select m.*,me.Value as EmployeeNo from Member m                            
left join Member_Extended_Property_Value me on me.Member_RecID = m.Member_RecID and me.Member_Extended_Property_RecID = 98) m on m.Member_RecID = t.Member_RecID                            
inner join Activity_Type at on at.Activity_Type_RecID = t.Activity_Type_RecID                            
left outer join holiday h on convert(varchar, h.Holiday_Date, 112) = convert(varchar, t.Date_Start, 112)                            
where --((t.Date_Start between '" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'--'" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'                            
--and at.Utilization_Flag = 0 and at.Inactive_Flag = 0 and at.Activity_Type_RecID = 7 and m.Member_Type_RecID != 5)) --and t.Member_RecID = @member_recid                            
(t.Date_Start between '" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'--'" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'                            
and m.Member_Type_RecID = 3  and at.Utilization_Flag = 1 and at.Inactive_Flag = 0)                            
Group by m.Member_RecID                            
,m.EmployeeNo                            
,m.First_Name + ' ' + m.Last_Name                            
--,t.Date_Start                            
--, at.Description                            
,at.Integration_Xref                            
,at.Xref_Work_Type                            
,case when h.Holiday_Date is not null then 1 else 0 end                            
union all                            
-- Övertid                            
Select                            
m.Member_RecID as MemberKey                            
,cast(m.EmployeeNo as int) EmployeeNo                            
,m.First_Name + ' ' + m.Last_Name as MemberName                
,'" + From.ToShortDateString() + @"'                
,'" + Tom.ToShortDateString() + @"'                            
,(case                             
when h.Holiday_Date is not null then 'Övertid'                            
when datepart(WEEKDAY, t.Date_Start) = 1 or datepart(WEEKDAY, t.Date_Start) = 7 then  'Övertid'                            
when t.Time_Start >= '18:00:00' then 'Övertid'                            
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' then 'Övertid'                            
when t.Time_End >= '18:00:00' and t.Time_Start <= '18:00:00' then 'Övertid'                            
else at.Description end) as ActivityType                            
,(case                             
when h.Holiday_Date is not null then 1                            
when datepart(WEEKDAY, t.Date_Start) = 1 or datepart(WEEKDAY, t.Date_Start) = 7 then  1                            
when t.Time_Start >= '18:00:00' then 1                            
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' then 1                           
when t.Time_End >= '18:00:00' and t.Time_Start <= '18:00:00' then 1                            
else 0 end) IsOverTime                            
,sum(cast((case                             
when h.Holiday_Date is not null then cast(datepart(MINUTE, t.Time_End - t.Time_Start) as decimal) / 60 + datepart(HOUR, t.Time_End - t.Time_Start)                            
when datepart(WEEKDAY, t.Date_Start) = 1 or datepart(WEEKDAY, t.Date_Start) = 7 then  cast(datepart(MINUTE, t.Time_End - t.Time_Start) as decimal) / 60 + datepart(HOUR, t.Time_End - t.Time_Start)                            
when t.Time_Start >= '18:00:00' then cast(datepart(MINUTE, t.Time_End - t.Time_Start) as decimal) / 60 + datepart(HOUR, t.Time_End - t.Time_Start)                            
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' and t.Time_End < '07:00:00' then cast(datepart(MINUTE, t.Time_End - t.Time_Start) as decimal) / 60 + datepart(HOUR, t.Time_End - t.Time_Start)                
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' and t.Time_End >= '07:00:00' then cast(datepart(MINUTE, ('07:00:00') - t.Time_Start) as decimal) / 60 + datepart(HOUR, ('07:00:00') - t.Time_Start)      
when t.Time_End >= '18:00:00' and t.Time_Start <= '18:00:00' then Cast(datepart(MINUTE, t.Time_End - '18:00:00') as decimal) / 60 + datepart(HOUR, t.Time_End - '18:00:00')                            
end) as numeric(5, 2))) as OverTime                            
,case when h.Holiday_Date is not null then 1 else 0 end as IsHoliday                            
,sum(cast((case                             
when h.Holiday_Date is not null then isnull(cast(datepart(MINUTE, t.Time_End - t.Time_Start) as decimal) / 60 + datepart(HOUR, t.Time_End - t.Time_Start), t.Hours_Actual)                            
when datepart(WEEKDAY, t.Date_Start) = 1 or datepart(WEEKDAY, t.Date_Start) = 7 then  isnull(cast(datepart(MINUTE, t.Time_End - t.Time_Start) as decimal) / 60 + datepart(HOUR, t.Time_End - t.Time_Start), t.Hours_Actual) 
when t.Time_Start >= '18:00:00' then isnull(cast(datepart(MINUTE, t.Time_End - t.Time_Start) as decimal) / 60 + datepart(HOUR, t.Time_End - t.Time_Start), t.Hours_Actual)                            
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' and t.Time_End < '07:00:00' then cast(datepart(MINUTE, t.Time_End - t.Time_Start) as decimal) / 60 + datepart(HOUR, t.Time_End - t.Time_Start)
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' and t.Time_End >= '07:00:00' then cast(datepart(MINUTE, ('07:00:00') - t.Time_Start) as decimal) / 60 + datepart(HOUR, ('07:00:00') - t.Time_Start) 
when t.Time_End >= '18:00:00' and t.Time_Start <= '18:00:00' then isnull(Cast(datepart(MINUTE, t.Time_End - '18:00:00') as decimal) / 60 + datepart(HOUR, t.Time_End - '18:00:00'), t.Hours_Actual)                            
else t.Hours_Actual end) as numeric(5, 2))) As Hours                             
,'315' LoneArtsNr                            
,'x315' Account                            
from Time_Entry t                            
inner join Time_Sheet ts on ts.Time_Sheet_RecID = t.Time_Sheet_RecID                           
inner join TE_Period tp on tp.TE_Period_RecID = ts.TE_Period_RecID                            
inner join (Select m.*, me.Value as EmployeeNo from Member m                            
left join Member_Extended_Property_Value me on me.Member_RecID = m.Member_RecID and me.Member_Extended_Property_RecID = 98) m on m.Member_RecID = t.Member_RecID                            
inner join Activity_Type at on at.Activity_Type_RecID = t.Activity_Type_RecID                            left outer join holiday h on convert(varchar, h.Holiday_Date, 112) = convert(varchar, t.Date_Start, 112)   
where((t.Date_Start between '" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'--'" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'                            
and at.Utilization_Flag = 1 and at.Inactive_Flag = 0 and m.Member_Type_RecID not in (5,6))                            
or(t.Date_Start between '" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'--'" + From.ToShortDateString() + @"' and '" + Tom.ToShortDateString() + @"'                            
and m.Member_Type_RecID = 3  and at.Utilization_Flag = 1 and at.Inactive_Flag = 0 and m.Member_Type_RecID not in (5,6)))                            
and(h.Holiday_Date is not null                           
or(datepart(WEEKDAY, t.Date_Start) = 1 or datepart(WEEKDAY, t.Date_Start) = 7)                            
OR t.Time_Start >= '18:00:00'                           
OR t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00'                            
or t.Time_End >= '18:00:00' and t.Time_Start <= '18:00:00'                            )                           
group by m.Member_RecID                            
,m.EmployeeNo                            
,m.First_Name + ' ' + m.Last_Name                            
--,t.Date_Start                            
,(case                            
when h.Holiday_Date is not null then 'Övertid'                           
when datepart(WEEKDAY, t.Date_Start) = 1 or datepart(WEEKDAY, t.Date_Start) = 7 then  'Övertid'                           
when t.Time_Start >= '18:00:00' then 'Övertid'                            
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' then 'Övertid'                            
when t.Time_End >= '18:00:00' and t.Time_Start <= '18:00:00' then 'Övertid'                            
else at.Description end)                             
,(case                            
when h.Holiday_Date is not null then 1                            
when datepart(WEEKDAY, t.Date_Start) = 1 or datepart(WEEKDAY, t.Date_Start) = 7 then  1                            
when t.Time_Start >= '18:00:00' then 1                            
when t.Time_Start <= '07:00:00' and t.Time_Start >= '00:00:00' then 1                           
when t.Time_End >= '18:00:00' and t.Time_Start <= '18:00:00' then 1                            
else 0 end)                             
,case when h.Holiday_Date is not null then 1 else 0 end                            
order by EmployeeNo, Date_Start, at.Description
";
            return select;


        }

        private void btnChangeSaveDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Flex";
            
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                txtSaveFileToDir.Text = fbd.SelectedPath;
            }
            else
            {
                txtSaveFileToDir.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Flex";
            }
        }

        private void btnCreateFile_Click(object sender, EventArgs e)
        {
            string file = txtSaveFileToDir.Text + @"\Lon_Acon_" + dFirstDayOfLastMonth.ToShortDateString().Replace("-", "") + "_" + 
                dLastDayOfLastMonth.ToShortDateString().Replace("-", "") + "_" +
                DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace(" ","") + ".DTA";
            StreamWriter sw = new StreamWriter(file);
            
            foreach(string lbi in lbCWData.Items)
            {
                sw.WriteLine(lbi);
            }
            sw.Close();
            ss1.Items.Clear();
            ss1.Items.Add("Skapat fil: " + file);
            ss1.ForeColor = Color.Green;
        }

        
    }
}
