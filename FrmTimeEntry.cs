//Author: Steven
//Purpose: This class controls sheduling

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Collections;
using System.Configuration;
using System.Data.OleDb;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Reflection;





namespace Control_Payroll
{
    public partial class FrmTimeEntry : Form
    {
        public FrmTimeEntry()
        {
            InitializeComponent();
        }

        UseDatabase useDB = new UseDatabase(GlobalClass.dbpathmet());
        OleDbDataReader dbReader = null;
        string query = "";

        string CurrEmpPayFreq = "";
        bool isReadOnly = false;
        string payFreq = "";
        string payday = "";
        bool scheduled = false;
        int schedID = 0;
        int EmployerSelectId = 0;


        private void FrmTimeEntry_Load(object sender, EventArgs e)
        {
            useDB.ConnectToDatabase();
            query = "SELECT EMPLOYER_NAME FROM EMPLOYER WHERE ACTIVE = true";
            dbReader = useDB.ExecuteQuery(query);
            dbReader.Read();

            if(dbReader.HasRows)
            {
                do
                {
                    if (!comboBox1.Items.Contains(dbReader["EMPLOYER_NAME"].ToString()))
                    {
                        comboBox1.Items.Add(dbReader["EMPLOYER_NAME"].ToString());
                    }
                }
                while (dbReader.Read()); 
            }
            if (EmployerSelectId == 0)
            {
                comboBox1.SelectedIndex = 0;
            }
            else
            {
                comboBox1.SelectedIndex = EmployerSelectId;
            }
        }

        private void btnEmpLoad_Click(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            List<int> IDList = new List<int>();
            if (comboBox1.SelectedIndex != -1)
            {
                query = "SELECT EMPLOYEE_ID FROM EMPEMP WHERE EMPLOYER_NAME = '" + comboBox1.Text + "' AND ACTIVE = true";
                dbReader = useDB.ExecuteQuery(query);
                dbReader.Read();

                if (dbReader.HasRows)
                {
                    do
                    {
                        if (!IDList.Contains(int.Parse(dbReader["EMPLOYEE_ID"].ToString())))
                        {
                            IDList.Add(int.Parse(dbReader["EMPLOYEE_ID"].ToString()));
                        }

                    } while (dbReader.Read());
                }



                IDList.Sort();

                for (int i = 0; i < IDList.Count; i++)
                {
                    if (!comboBox2.Items.Contains(IDList[i]))
                    {
                        query = "SELECT FIRSTNAME, LASTNAME FROM EMPLOYEE WHERE ID = " + IDList[i] + "";
                        dbReader = useDB.ExecuteQuery(query);
                        dbReader.Read();
                        comboBox2.Items.Add("" + IDList[i] + " - " + dbReader["FIRSTNAME"].ToString() + " " + dbReader["LASTNAME"].ToString());
                    }
                }

                comboBox2.SelectedIndex = 0;
                comboBox2_SelectedIndexChanged(this, null);
            }
            else
            {
                MessageBox.Show("Please select a client from the list", "Error");
            }
        }

        private void btnexit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (System.Windows.Forms.Application.OpenForms["FrmTimeEntryHourly"] != null)
            {
                (System.Windows.Forms.Application.OpenForms["FrmTimeEntryHourly"] as FrmTimeEntryHourly).Close();
            }
            scheduled = false;
            int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
            query = "SELECT * FROM PAY_RUN WHERE EMPLOYEE_ID = " + index + " AND DATE_SCHEDULED = '" + DateTime.Today.ToShortDateString() + "' AND RELEASED = false";
            dbReader = useDB.ExecuteQuery(query);
            dbReader.Read();
            
            if (dbReader.HasRows)
            {
                label26.Visible = true;
                scheduled = true;
                schedID = int.Parse(dbReader["ID"].ToString());
                loadExistingSchedule(schedID, dbReader["PAY_FREQ"].ToString());
                payFreq = dbReader["PAY_FREQ"].ToString();

                query = "SELECT * FROM EMPLOYEE WHERE ID = " + index + "";
                dbReader = useDB.ExecuteQuery(query);
                dbReader.Read();

                if (dbReader.HasRows)
                {
                    lblEmpID.Text = dbReader["ID_NUMBER"].ToString();
                    lblEmpName.Text = dbReader["FIRSTNAME"].ToString() + " " + dbReader["LASTNAME"].ToString();
                    lblBankAccountNum.Text = dbReader["BANK_ACCNUM"].ToString();
                    lblBankName.Text = dbReader["BANK_NAME"].ToString();

                    lblDOB.Text = "N/A";

                    string DOBYear = lblEmpID.Text.Substring(0, 2);
                    string DOBMonth = lblEmpID.Text.Substring(2, 2);
                    string DOBDay = lblEmpID.Text.Substring(4, 2);

                    if (int.Parse("20" + DOBYear) < DateTime.Today.Year)
                    {
                        DOBYear = "20" + DOBYear;
                    }
                    else
                    {
                        DOBYear = "19" + DOBYear;
                    }

                    lblDOB.Text = DOBDay + "/" + DOBMonth + "/" + DOBYear;
                }
                else
                {
                    MessageBox.Show("No info on Employee Found");
                }

                query = "SELECT * FROM EMPEMP WHERE EMPLOYEE_ID = " + index + " AND ACTIVE = true";
                dbReader = useDB.ExecuteQuery(query);
                dbReader.Read();

                if (dbReader.HasRows)
                {
                    txtRate.Text = double.Parse(dbReader["PAYRATE"].ToString()).ToString("#,##0.00");
                    payFreq = dbReader["PAY_FREQ"].ToString();
                    payday = dbReader["PAYDAY"].ToString();

                    if (dbReader["PAY_PERRATE"].ToString().ToUpper() == "HOUR")
                    {
                        lblPerRate.Text = "Per Hour";
                    }
                    else if (dbReader["PAY_PERRATE"].ToString().ToUpper() == "DAY")
                    {
                        lblPerRate.Text = "Per Day";
                    }
                    else if (dbReader["PAY_PERRATE"].ToString().ToUpper() == "WEEK")
                    {
                        lblPerRate.Text = "Per Week";
                    }
                    else if (dbReader["PAY_PERRATE"].ToString().ToUpper() == "BIWEEK")
                    {
                        lblPerRate.Text = "Per Biweek";
                    }
                    else
                    {
                        lblPerRate.Text = "Per Month";
                    }
                }
            }
            else
            {
                scheduled = false;
                label26.Visible = false;

                query = "SELECT * FROM EMPLOYEE WHERE ID = " + index + "";
                dbReader = useDB.ExecuteQuery(query);
                dbReader.Read();

                if (dbReader.HasRows)
                {
                    lblEmpID.Text = dbReader["ID_NUMBER"].ToString();
                    lblEmpName.Text = dbReader["FIRSTNAME"].ToString() + " " + dbReader["LASTNAME"].ToString();
                    lblBankAccountNum.Text = dbReader["BANK_ACCNUM"].ToString();
                    lblBankName.Text = dbReader["BANK_NAME"].ToString();

                    lblDOB.Text = "N/A";

                    string DOBYear = lblEmpID.Text.Substring(0, 2);
                    string DOBMonth = lblEmpID.Text.Substring(2, 2);
                    string DOBDay = lblEmpID.Text.Substring(4, 2);

                    if (int.Parse("20" + DOBYear) < DateTime.Today.Year)
                    {
                        DOBYear = "20" + DOBYear;
                    }
                    else
                    {
                        DOBYear = "19" + DOBYear;
                    }

                    lblDOB.Text = DOBDay + "/" + DOBMonth + "/" + DOBYear;
                }
                else
                {
                    MessageBox.Show("No info on Employee Found");
                }

                query = "SELECT * FROM EMPEMP WHERE EMPLOYEE_ID = " + index + " AND ACTIVE = true";
                dbReader = useDB.ExecuteQuery(query);
                dbReader.Read();

                if (dbReader.HasRows)
                {
                    txtRate.Text = double.Parse(dbReader["PAYRATE"].ToString()).ToString("#,##0.00");
                    payFreq = dbReader["PAY_FREQ"].ToString();
                    payday = dbReader["PAYDAY"].ToString();

                    if (dbReader["PAY_PERRATE"].ToString().ToUpper() == "HOUR")
                    {
                        lblPerRate.Text = "Per Hour";
                    }
                    else if (dbReader["PAY_PERRATE"].ToString().ToUpper() == "DAY")
                    {
                        lblPerRate.Text = "Per Day";
                    }
                    else if (dbReader["PAY_PERRATE"].ToString().ToUpper() == "WEEK")
                    {
                        lblPerRate.Text = "Per Week";
                    }
                    else if (dbReader["PAY_PERRATE"].ToString().ToUpper() == "BIWEEK")
                    {
                        lblPerRate.Text = "Per Biweek";
                    }
                    else
                    {
                        lblPerRate.Text = "Per Month";
                    }

                    double OTRate = 0.00;

                    hourlyEntryUncollapsed1.Visible = true;
                    dailyEntryUncollapsed1.Visible = true;
                    weeklyEntryUncollapsed1.Visible = true;
                    biweeklyEntryUncollapsed1.Visible = true;
                    monthlyEntryUncollapsed1.Visible = true;

                    if (payFreq.ToUpper() == "HOURLY")
                    {
                        hourlyEntryUncollapsed1.loadDefaults(txtRate.Text);
                    }
                    else if (payFreq.ToUpper() == "DAILY")
                    {
                        dailyEntryUncollapsed1.loadDefaults(txtRate.Text);
                    }
                    else if (payFreq.ToUpper() == "WEEKLY")
                    {
                        if (lblPerRate.Text == "Per Week")
                        {
                            OTRate = (double.Parse(txtRate.Text) / 45) * 1.5;
                        }
                        else if (lblPerRate.Text == "Per Month")
                        {
                            OTRate = (double.Parse(txtRate.Text) / 195) * 1.5;
                        }
                        weeklyEntryUncollapsed1.loadDefaults(txtRate.Text, OTRate);
                    }
                    else if (payFreq.ToUpper() == "BIWEEKLY")
                    {
                        if (lblPerRate.Text == "Per Biweek")
                        {
                            OTRate = (double.Parse(txtRate.Text) / 90) * 1.5;
                        }
                        else if (lblPerRate.Text == "Per Month")
                        {
                            OTRate = (double.Parse(txtRate.Text) / 195) * 1.5;
                        }
                        biweeklyEntryUncollapsed1.loadDefaults(txtRate.Text, OTRate);
                    }
                    else
                    {
                        OTRate = (double.Parse(txtRate.Text) / 195) * 1.5;
                        monthlyEntryUncollapsed1.loadDefaults(txtRate.Text, OTRate);
                    }

                }

                hourlyEntryCollapsed1.Visible = true;
                dailyEntryCollapsed1.Visible = true;
                weeklyEntryCollapsed1.Visible = true;
                biweeklyEntryCollapsed1.Visible = true;
                monthlyEntryCollapsed1.Visible = true;

                hourlyEntryUncollapsed1.Visible = false;
                dailyEntryUncollapsed1.Visible = false;
                weeklyEntryUncollapsed1.Visible = false;
                biweeklyEntryUncollapsed1.Visible = false;
                monthlyEntryUncollapsed1.Visible = false;

                if (scheduled == false)
                {
                    additionsUncollapsed1.Visible = true;
                    deductionsUncollapsed1.Visible = true;
                    additionsUncollapsed1.loadAdditions(index);
                    deductionsUncollapsed1.loadDeductions(index);
                    additionsUncollapsed1.Visible = false;
                    deductionsUncollapsed1.Visible = false;
                }
            }
            #region PayFrequencyControl
            if (payFreq.ToUpper() == "HOURLY")
            {
                hourlyEntryCollapsed1.ReadOnlyLoad(false);
                dailyEntryCollapsed1.ReadOnlyLoad(true);
                weeklyEntryCollapsed1.ReadOnlyLoad(true);
                biweeklyEntryCollapsed1.ReadOnlyLoad(true);
                monthlyEntryCollapsed1.ReadOnlyLoad(true);
            }
            else if (payFreq.ToUpper() == "DAILY")
            {
                hourlyEntryCollapsed1.ReadOnlyLoad(true);
                dailyEntryCollapsed1.ReadOnlyLoad(false);
                weeklyEntryCollapsed1.ReadOnlyLoad(true);
                biweeklyEntryCollapsed1.ReadOnlyLoad(true);
                monthlyEntryCollapsed1.ReadOnlyLoad(true);
                //monthly
            }
            else if (payFreq.ToUpper() == "WEEKLY")
            {
                hourlyEntryCollapsed1.ReadOnlyLoad(true);
                dailyEntryCollapsed1.ReadOnlyLoad(true);
                weeklyEntryCollapsed1.ReadOnlyLoad(false);
                biweeklyEntryCollapsed1.ReadOnlyLoad(true);
                monthlyEntryCollapsed1.ReadOnlyLoad(true);
                //monthly
            }
            else if (payFreq.ToUpper() == "BIWEEKLY")
            {
                hourlyEntryCollapsed1.ReadOnlyLoad(true);
                dailyEntryCollapsed1.ReadOnlyLoad(true);
                weeklyEntryCollapsed1.ReadOnlyLoad(true);
                biweeklyEntryCollapsed1.ReadOnlyLoad(false);
                monthlyEntryCollapsed1.ReadOnlyLoad(true);
            }
            else
            {
                hourlyEntryCollapsed1.ReadOnlyLoad(true);
                dailyEntryCollapsed1.ReadOnlyLoad(true);
                weeklyEntryCollapsed1.ReadOnlyLoad(true);
                biweeklyEntryCollapsed1.ReadOnlyLoad(true);
                monthlyEntryCollapsed1.ReadOnlyLoad(false);

            }
            #endregion
            }




        public void changeControlVisibility(string PayFrequency)
        {
            if (PayFrequency == "hourly")
            {
                if (hourlyEntryCollapsed1.Visible == true)
                {
                    hourlyEntryUncollapsed1.Visible = true;
                    hourlyEntryCollapsed1.Visible = false;
                }
                else
                {
                    hourlyEntryUncollapsed1.Visible = false;
                    hourlyEntryCollapsed1.Visible = true;
                }
            }
            else if (PayFrequency == "daily")
            {
                if (dailyEntryCollapsed1.Visible == true)
                {
                    dailyEntryUncollapsed1.Visible = true;
                    dailyEntryCollapsed1.Visible = false;
                }
                else
                {
                    dailyEntryUncollapsed1.Visible = false;
                    dailyEntryCollapsed1.Visible = true;
                }
            }
            else if (PayFrequency == "weekly")
            {
                if (weeklyEntryCollapsed1.Visible == true)
                {
                    weeklyEntryUncollapsed1.Visible = true;
                    weeklyEntryCollapsed1.Visible = false;
                }
                else
                {
                    weeklyEntryUncollapsed1.Visible = false;
                    weeklyEntryCollapsed1.Visible = true;
                }
            }
            else if (PayFrequency == "biweekly")
            {
                if (biweeklyEntryCollapsed1.Visible == true)
                {
                    biweeklyEntryUncollapsed1.Visible = true;
                    biweeklyEntryCollapsed1.Visible = false;
                }
                else
                {
                    biweeklyEntryUncollapsed1.Visible = false;
                    biweeklyEntryCollapsed1.Visible = true;
                }
            }
            else if (PayFrequency == "monthly")
            {
                if (monthlyEntryCollapsed1.Visible == true)
                {
                    monthlyEntryUncollapsed1.Visible = true;
                    monthlyEntryCollapsed1.Visible = false;
                }
                else
                {
                    monthlyEntryUncollapsed1.Visible = false;
                    monthlyEntryCollapsed1.Visible = true;
                }
            }
            else if (PayFrequency == "additions")
            {
                if (additionsCollapsed1.Visible == true)
                {
                    additionsUncollapsed1.Visible = true;
                    additionsCollapsed1.Visible = false;
                }
                else
                {
                    additionsUncollapsed1.Visible = false;
                    additionsCollapsed1.Visible = true;
                }
            }
            else if (PayFrequency == "deductions")
            {
                if (deductionsCollapsed1.Visible == true)
                {
                    deductionsUncollapsed1.Visible = true;
                    deductionsCollapsed1.Visible = false;
                }
                else
                {
                    deductionsUncollapsed1.Visible = false;
                    deductionsCollapsed1.Visible = true;
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (System.Windows.Forms.Application.OpenForms["FrmTimeEntryHourly"] as FrmTimeEntryHourly != null)
            {
                (System.Windows.Forms.Application.OpenForms["FrmTimeEntryHourly"] as FrmTimeEntryHourly).Close();
            }
            EmployerSelectId = comboBox1.SelectedIndex;
            btnEmpLoad_Click(sender, null);
        }

        public void LoadHourlyTimeEntry()
        {
            FrmTimeEntryHourly FTEH = new FrmTimeEntryHourly();
            FTEH.CreateTimeTable(dateTimePicker2.Value, dateTimePicker1.Value);

            FTEH.Show();
        }

        private void Additionsbtn_Click(object sender, EventArgs e)
        {
            int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
            AdditionsPage addPage = new AdditionsPage();
            addPage.loadAdditionsPage(index, lblEmpName.Text, lblEmpID.Text);
            addPage.Show();
        }

        private void activeDeducBTN_Click(object sender, EventArgs e)
        {
            int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
            DeducPage deducPage = new DeducPage();
            deducPage.loadDeducPage(index, lblEmpName.Text, lblEmpID.Text);
            deducPage.Show();
            /*deducSetup dedSetup = new deducSetup();
            dedSetup.loadDeduction(int.Parse(comboBox2.SelectedItem.ToString()), lblEmpName.Text, lblEmpID.Text);
            dedSetup.Show();*/
        }

        public void UpdateDeductions()
        {
            if (deductionsUncollapsed1.Visible == false)
            {
                deductionsUncollapsed1.Visible = true;
                deductionsCollapsed1.Visible = false;
                int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
                deductionsUncollapsed1.loadDeductions(index);
                deductionsUncollapsed1.Visible = false;
                deductionsCollapsed1.Visible = true;
            }
            else
            {
                deductionsCollapsed1.Visible = false;
                int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
                deductionsUncollapsed1.loadDeductions(index);
            }
        }

        public void UpdateAdditions()
        {
            if (additionsUncollapsed1.Visible == false)
            {
                additionsUncollapsed1.Visible = true;
                additionsCollapsed1.Visible = false;
                int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
                additionsUncollapsed1.loadAdditions(index);
                additionsUncollapsed1.Visible = false;
                additionsCollapsed1.Visible = true;
            }
            else
            {
                additionsCollapsed1.Visible = false;
                int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
                additionsUncollapsed1.loadAdditions(index);
            }
        }

        public void UpdateAdditionsTable()
        {
            int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
            additionsUncollapsed1.AdditionTableUpdate(index);
        }

        public void UpdateDeductionsTable()
        {
            int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
            deductionsUncollapsed1.DeductionsTableUpdate();
        }

        private void BtnRunWork_Click(object sender, EventArgs e)
        {
            int index = comboBox2.SelectedIndex;
            if (scheduled == true)
            {
                DialogResult result = MessageBox.Show("Employee already scheduled for release, would you like to update the schedule?", "Error", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    UpdateSchedule(schedID);
                }
                else
                {
                    comboBox2.SelectedIndex = index;
                }
                
            }
            else
            {
                InsertSchedule();
            }
        }

        private void InsertSchedule()
        {
            List<string> DEDONCE_DESC = new List<string>();
            List<double> DEDONCE_AMT = new List<double>();
            List<string> ALLONCE_DESC = new List<string>();
            List<double> ALLONCE_AMT = new List<double>();
            int EMPLOYER_ID = 0;
            string EMPLOYER_NAME = "";
            int EMPLOYEE_ID = 0;
            string EMPLOYEE_NAME = "";
            double PAY_AMT = 0.00;
            int EMPEMP_ID = 0;
            string PAY_FREQ = "";
            string PAY_PERIODSTART = "";
            string PAY_SCHEDDATE = "";
            bool RELEASED = false;
            bool PAYSLIPS_RAN = false;
            string NT = "";
            string OT = "";
            string DT = "";
            string SUNDAY = "";
            string BANK_ACC_NUM = "";
            string BANK_NAME = "";
            string AGE = "";
            string OTPERTYPE = "";
            double OTPER = 0;
            string PAYDAY = "";
            string DATE_SCHEDULED = "";

            //Setting info

            query = "SELECT ID FROM EMPLOYER WHERE EMPLOYER_NAME = '" + comboBox1.Text + "'";
            dbReader = useDB.ExecuteQuery(query);
            dbReader.Read();

            EMPLOYER_ID = int.Parse(dbReader["ID"].ToString());
            EMPLOYER_NAME = comboBox1.Text;
            int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
            EMPLOYEE_ID = index;
            EMPLOYEE_NAME = lblEmpName.Text;
            PAY_AMT = double.Parse(txtRate.Text);

            query = "SELECT ID FROM EMPEMP WHERE EMPLOYEE_ID = " + index + " AND ACTIVE = true";
            dbReader = useDB.ExecuteQuery(query);
            dbReader.Read();

            EMPEMP_ID = int.Parse(dbReader["ID"].ToString());

            query = "SELECT PAY_FREQ FROM EMPEMP WHERE EMPLOYEE_ID = " + index + " AND ACTIVE = true";
            dbReader = useDB.ExecuteQuery(query);
            dbReader.Read();

            PAY_FREQ = dbReader["PAY_FREQ"].ToString();
            PAY_PERIODSTART = dateTimePicker2.Value.ToShortDateString();
            PAY_SCHEDDATE = dateTimePicker1.Value.ToShortDateString();
            RELEASED = false;

            for (int i = 1; i < 9; i++)
            {
                DEDONCE_DESC.Add(deductionsUncollapsed1.GetDeductionsDescForSchedule(i));
                DEDONCE_AMT.Add(deductionsUncollapsed1.GetDeductionsAmtForSchedule(i));
            }

            for (int i = 1; i < 5; i++)
            {
                ALLONCE_DESC.Add(additionsUncollapsed1.GetAdditionsDescForSchedule(i));
                ALLONCE_AMT.Add(additionsUncollapsed1.GetAdditionsAmtForSchedule(i));
            }

            PAYSLIPS_RAN = false;

            if (PAY_FREQ.ToUpper() == "HOURLY")
            {
                PAY_AMT = double.Parse(txtRate.Text);
                NT = hourlyEntryUncollapsed1.ReturnNormalTime();
                OT = hourlyEntryUncollapsed1.ReturnOverTime();
                DT = hourlyEntryUncollapsed1.ReturnDoubleTime();
                SUNDAY = hourlyEntryUncollapsed1.ReturnSundayTime();
            }
            else if (PAY_FREQ.ToUpper() == "DAILY")
            {
                PAY_AMT = double.Parse(txtRate.Text);
                NT = dailyEntryUncollapsed1.ReturnNormalDays();
                DT = dailyEntryUncollapsed1.ReturnDoubleDays();
                OT = dailyEntryUncollapsed1.ReturnOverDays();
                OTPER = dailyEntryUncollapsed1.ReturnOTPER();
                OTPERTYPE = dailyEntryUncollapsed1.ReturnOTPERTYPE();
            }
            else if (PAY_FREQ.ToUpper() == "WEEKLY")
            {
                PAY_AMT = weeklyEntryUncollapsed1.returnPAY_AMT();
                OT = weeklyEntryUncollapsed1.returnOverTime();
                OTPERTYPE = weeklyEntryUncollapsed1.returnOTPERTYPE();
                OTPER = weeklyEntryUncollapsed1.returnOTPER();
            }
            else if (PAY_FREQ.ToUpper() == "BIWEEKLY")
            {
                PAY_AMT = biweeklyEntryUncollapsed1.returnPAY_AMT();
                OT = biweeklyEntryUncollapsed1.returnOverTime();
                OTPERTYPE = biweeklyEntryUncollapsed1.returnOTPERTYPE();
                OTPER = biweeklyEntryUncollapsed1.returnOTPER();
            }
            else
            {
                PAY_AMT = monthlyEntryUncollapsed1.returnPAY_AMT();
                OT = monthlyEntryUncollapsed1.returnOverTime();
                OTPERTYPE = monthlyEntryUncollapsed1.returnOTPERTYPE();
                OTPER = monthlyEntryUncollapsed1.returnOTPER();
            }

            BANK_ACC_NUM = lblBankAccountNum.Text;
            BANK_NAME = lblBankName.Text;

            query = "SELECT PAYDAY FROM EMPEMP WHERE EMPLOYEE_ID = " + index + " AND ACTIVE = true";
            dbReader = useDB.ExecuteQuery(query);
            dbReader.Read();

            PAYDAY = dbReader["PAYDAY"].ToString();
            DATE_SCHEDULED = DateTime.Today.ToShortDateString();

            AGE = (DateTime.Today.Year - DateTime.Parse(lblDOB.Text).Year).ToString();

            query = "INSERT INTO PAY_RUN(EMPLOYER_ID, EMPLOYER_NAME, EMPLOYEE_ID, EMPLOYEE_NAME, PAY_AMT, EMPEMP_ID, PAY_FREQ, "
                    + "PAY_PERIODSTART, PAY_SCHEDDATE, RELEASED, PAY_DEDONCE1TXT, PAY_DEDONCE1AMT, PAY_DEDONCE2TXT, PAY_DEDONCE2AMT, PAY_DEDONCE3TXT, "
                    + "PAY_DEDONCE3AMT, PAY_DEDONCE4TXT, PAY_DEDONCE4AMT, PAY_DEDONCE5TXT, PAY_DEDONCE5AMT, PAY_DEDONCE6TXT, PAY_DEDONCE6AMT, PAY_DEDONCE7TXT, "
                    + "PAY_DEDONCE7AMT, PAY_DEDONCE8TXT, PAY_DEDONCE8AMT, PAY_ALLONCE1TXT, PAY_ALLONCE1AMT, PAY_ALLONCE2TXT, PAY_ALLONCE2AMT, PAY_ALLONCE3TXT, "
                    + "PAY_ALLONCE3AMT, PAY_ALLONCE4TXT, PAY_ALLONCE4AMT, PAYSLIPS_RAN, NT, OT, DT, SUNDAY, BANK_ACC_NUM, BANK_NAME, AGE, OTPERTYPE, OTPER, "
                    + "PAYDAY, DATE_SCHEDULED) VALUES(" + EMPLOYER_ID + ", '" + EMPLOYER_NAME + "', " + EMPLOYEE_ID + ", '" + EMPLOYEE_NAME + "', " + PAY_AMT + ", " + EMPEMP_ID + ", '" + PAY_FREQ + "', '"
                    + PAY_PERIODSTART + "', '" + PAY_SCHEDDATE + "', " + RELEASED + ", '" + DEDONCE_DESC[0] + "', " + DEDONCE_AMT[0] + ", '" + DEDONCE_DESC[1] + "', " + DEDONCE_AMT[1] + ", '" + DEDONCE_DESC[2] + "', "
                    + DEDONCE_AMT[2] + ", '" + DEDONCE_DESC[3] + "', " + DEDONCE_AMT[3] + ", '" + DEDONCE_DESC[4] + "', " + DEDONCE_AMT[4] + ", '" + DEDONCE_DESC[5] + "', " + DEDONCE_AMT[5] + ", '" + DEDONCE_DESC[6] + "', "
                    + DEDONCE_AMT[6] + ", '" + DEDONCE_DESC[7] + "', " + DEDONCE_AMT[7] + ", '" + ALLONCE_DESC[0] + "', " + ALLONCE_AMT[0] + ", '" + ALLONCE_DESC[1] + "', " + ALLONCE_AMT[1] + ", '" + ALLONCE_DESC[2] + "', "
                    + ALLONCE_AMT[2] + ", '" + ALLONCE_DESC[3] + "', " + ALLONCE_AMT[3] + ", " + PAYSLIPS_RAN + ", '" + NT + "', '" + OT + "', '" + DT + "', '" + SUNDAY + "', '" + BANK_ACC_NUM + "', '" + BANK_NAME + "', '" + AGE + "', '" + OTPERTYPE + "', " + OTPER + ", '"
                    + PAYDAY + "', '" + DATE_SCHEDULED + "')";

            useDB.ExecuteCommand(query);

            /*UpdateAdditionsTable();
            UpdateDeductionsTable();*/

            MessageBox.Show("Employee successfully Scheduled", "Success");

            int cmbIndex = comboBox2.SelectedIndex;
            comboBox2.SelectedIndex = cmbIndex;
            comboBox2_SelectedIndexChanged(this, null);
        }

        private void UpdateSchedule(int schedID)
        {
            List<string> DEDONCE_DESC = new List<string>();
            List<double> DEDONCE_AMT = new List<double>();
            List<string> ALLONCE_DESC = new List<string>();
            List<double> ALLONCE_AMT = new List<double>();

            double PAY_AMT = 0.00;
            string PAY_FREQ = "";
            string PAY_PERIODSTART = "";
            string PAY_SCHEDDATE = "";
            string NT = "";
            string OT = "";
            string DT = "";
            string SUNDAY = "";
            string OTPERTYPE = "";
            double OTPER = 0;

            //Setting info

            query = "SELECT ID FROM EMPLOYER WHERE EMPLOYER_NAME = '" + comboBox1.Text + "'";
            dbReader = useDB.ExecuteQuery(query);
            dbReader.Read();
            int index = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
            PAY_AMT = double.Parse(txtRate.Text);

            query = "SELECT PAY_FREQ FROM EMPEMP WHERE EMPLOYEE_ID = " + index + " AND ACTIVE = true";
            dbReader = useDB.ExecuteQuery(query);
            dbReader.Read();

            PAY_FREQ = dbReader["PAY_FREQ"].ToString();
            
            PAY_PERIODSTART = dateTimePicker2.Value.ToShortDateString();
            PAY_SCHEDDATE = dateTimePicker1.Value.ToShortDateString();

            for (int i = 1; i < 9; i++)
            {
                DEDONCE_DESC.Add(deductionsUncollapsed1.GetDeductionsDescForSchedule(i));
                DEDONCE_AMT.Add(deductionsUncollapsed1.GetDeductionsAmtForSchedule(i));
            }

            for (int i = 1; i < 5; i++)
            {
                ALLONCE_DESC.Add(additionsUncollapsed1.GetAdditionsDescForSchedule(i));
                ALLONCE_AMT.Add(additionsUncollapsed1.GetAdditionsAmtForSchedule(i));
            }

            if (PAY_FREQ.ToUpper() == "HOURLY")
            {
                PAY_AMT = double.Parse(txtRate.Text);
                NT = hourlyEntryUncollapsed1.ReturnNormalTime();
                OT = hourlyEntryUncollapsed1.ReturnOverTime();
                DT = hourlyEntryUncollapsed1.ReturnDoubleTime();
                SUNDAY = hourlyEntryUncollapsed1.ReturnSundayTime();
            }
            else if (PAY_FREQ.ToUpper() == "DAILY")
            {
                PAY_AMT = double.Parse(txtRate.Text);
                NT = dailyEntryUncollapsed1.ReturnNormalDays();
                DT = dailyEntryUncollapsed1.ReturnDoubleDays();
                OT = dailyEntryUncollapsed1.ReturnOverDays();
                OTPER = dailyEntryUncollapsed1.ReturnOTPER();
                OTPERTYPE = dailyEntryUncollapsed1.ReturnOTPERTYPE();
            }
            else if (PAY_FREQ.ToUpper() == "WEEKLY")
            {
                PAY_AMT = weeklyEntryUncollapsed1.returnPAY_AMT();
                OT = weeklyEntryUncollapsed1.returnOverTime();
                OTPERTYPE = weeklyEntryUncollapsed1.returnOTPERTYPE();
                OTPER = weeklyEntryUncollapsed1.returnOTPER();
            }
            else if (PAY_FREQ.ToUpper() == "BIWEEKLY")
            {
                PAY_AMT = biweeklyEntryUncollapsed1.returnPAY_AMT();
                OT = biweeklyEntryUncollapsed1.returnOverTime();
                OTPERTYPE = biweeklyEntryUncollapsed1.returnOTPERTYPE();
                OTPER = biweeklyEntryUncollapsed1.returnOTPER();
            }
            else
            {
                PAY_AMT = monthlyEntryUncollapsed1.returnPAY_AMT();
                OT = monthlyEntryUncollapsed1.returnOverTime();
                OTPERTYPE = monthlyEntryUncollapsed1.returnOTPERTYPE();
                OTPER = monthlyEntryUncollapsed1.returnOTPER();
            }

            query = "UPDATE PAY_RUN SET PAY_AMT = " + PAY_AMT + " WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            query = "UPDATE PAY_RUN SET PAY_PERIODSTART = '" + PAY_PERIODSTART + "' WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            query = "UPDATE PAY_RUN SET PAY_SCHEDDATE = '" + PAY_SCHEDDATE + "' WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            //UPDATE DEDUCTION

            for (int i = 0; i < 8; i++)
            {
                query = "UPDATE PAY_RUN SET PAY_DEDONCE" + (i + 1) + "TXT = '" + DEDONCE_DESC[i] + "' WHERE ID = " + schedID + "";
                useDB.ExecuteCommand(query);

                query = "UPDATE PAY_RUN SET PAY_DEDONCE" + (i + 1) + "AMT = " + DEDONCE_AMT[i] + " WHERE ID = " + schedID + "";
                useDB.ExecuteCommand(query);
            }

            //UPDATE ADDITIONS
            for (int i = 0; i < 4; i++)
            {
                query = "UPDATE PAY_RUN SET PAY_ALLONCE" + (i + 1) + "TXT = '" + ALLONCE_DESC[i] + "' WHERE ID = " + schedID + "";
                useDB.ExecuteCommand(query);

                query = "UPDATE PAY_RUN SET PAY_ALLONCE" + (i + 1) + "AMT = " + ALLONCE_AMT[i] + " WHERE ID = " + schedID + "";
                useDB.ExecuteCommand(query);
            }

            query = "UPDATE PAY_RUN SET NT = '" + NT + "' WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            query = "UPDATE PAY_RUN SET OT = '" + OT + "' WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            query = "UPDATE PAY_RUN SET DT = '" + DT + "' WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            query = "UPDATE PAY_RUN SET SUNDAY = '" + SUNDAY + "' WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            query = "UPDATE PAY_RUN SET OTPERTYPE = '" + OTPERTYPE + "' WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            query = "UPDATE PAY_RUN SET OTPER = " + OTPER + " WHERE ID = " + schedID + "";
            useDB.ExecuteCommand(query);

            MessageBox.Show("Employee Schedule Updated", "Success");
            
            int cmbIndex = comboBox2.SelectedIndex;
            comboBox2.SelectedIndex = cmbIndex;

        }

        private void loadExistingSchedule(int schedID, string payFreq)
        {
            if (payFreq.ToUpper() == "HOURLY")
            {
                hourlyEntryUncollapsed1.loadExistingHourly(schedID);
                additionsUncollapsed1.loadExistingAdditions(schedID);
                deductionsUncollapsed1.loadExistingDeductions(schedID);
            }
            else if (payFreq.ToUpper() == "DAILY")
            {
                dailyEntryUncollapsed1.loadExistingDaily(schedID);
                additionsUncollapsed1.loadExistingAdditions(schedID);
                deductionsUncollapsed1.loadExistingDeductions(schedID);
            }
            else if (payFreq.ToUpper() == "WEEKLY")
            {
                weeklyEntryUncollapsed1.loadExistingWeekly(schedID);
                additionsUncollapsed1.loadExistingAdditions(schedID);
                deductionsUncollapsed1.loadExistingDeductions(schedID);
            }
            else if (payFreq.ToUpper() == "BIWEEKLY")
            {
                biweeklyEntryUncollapsed1.loadExistingBiweekly(schedID);
                additionsUncollapsed1.loadExistingAdditions(schedID);
                deductionsUncollapsed1.loadExistingDeductions(schedID);
            }
            else
            {
                monthlyEntryUncollapsed1.loadExistingMonthly(schedID);
                additionsUncollapsed1.loadExistingAdditions(schedID);
                deductionsUncollapsed1.loadExistingDeductions(schedID);
            }
        }

        public void UpdateHourlyEntry(double NT, double OT, double DT, double Sunday)
        {
            hourlyEntryUncollapsed1.ReturnTime(NT, OT, DT, Sunday);
        }

        public void InsertLeavePayout(decimal normLeave, decimal doubleLeave)
        {
            if (normLeave > 0)
            {
                int EMPLOYEE_ID = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
                int UID = 0;
                query = "SELECT TOP 1 UID FROM ALLOWANCES ORDER BY ID DESC";
                dbReader = useDB.ExecuteQuery(query);
                dbReader.Read();

                if (dbReader.HasRows)
                {
                    UID = int.Parse(dbReader["UID"].ToString()) + 1;
                }
                else
                {
                    UID = 1;
                }
                int ALLOWANCE_CODE = 4;
                string ALLOWANCE_DESC = "Annual Leave";
                decimal ALLOWANCE_AMT = decimal.Parse(txtRate.Text) * normLeave;
                bool PERMANENT = false;
                bool ONCEOFF = true;
                bool ACTIVE = true;

                query = "INSERT INTO ALLOWANCES(EMPLOYEE_ID, UID, ALLOWANCE_CODE, ALLOWANCE_DESC, ALLOWANCE_AMT, PERMANENT, ONCEOFF, ACTIVE) VALUES(" + EMPLOYEE_ID + ", " + UID + ", " + ALLOWANCE_CODE + ", '" + ALLOWANCE_DESC + "', " + ALLOWANCE_AMT + ", " + PERMANENT + ", " + ONCEOFF + ", " + ACTIVE + ")";
                useDB.ExecuteCommand(query);
            }

            if (doubleLeave > 0)
            {
                int EMPLOYEE_ID = int.Parse(System.Text.RegularExpressions.Regex.Replace(comboBox2.Text, "[^0-9]+", string.Empty));
                int UID = 0;
                query = "SELECT TOP 1 UID FROM ALLOWANCES ORDER BY ID DESC";
                dbReader = useDB.ExecuteQuery(query);
                dbReader.Read();

                if (dbReader.HasRows)
                {
                    UID = int.Parse(dbReader["UID"].ToString()) + 1;
                }
                else
                {
                    UID = 1;
                }
                int ALLOWANCE_CODE = 4;
                string ALLOWANCE_DESC = "Annual Leave Public";
                decimal ALLOWANCE_AMT = (decimal.Parse(txtRate.Text) * 2) * doubleLeave;
                bool PERMANENT = false;
                bool ONCEOFF = true;
                bool ACTIVE = true;

                query = "INSERT INTO ALLOWANCES(EMPLOYEE_ID, UID, ALLOWANCE_CODE, ALLOWANCE_DESC, ALLOWANCE_AMT, PERMANENT, ONCEOFF, ACTIVE) VALUES(" + EMPLOYEE_ID + ", " + UID + ", " + ALLOWANCE_CODE + ", '" + ALLOWANCE_DESC + "', " + ALLOWANCE_AMT + ", " + PERMANENT + ", " + ONCEOFF + ", " + ACTIVE + ")";
                useDB.ExecuteCommand(query);
            }

            UpdateAdditions();
        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            comboBox2_SelectedIndexChanged(this, null);
        }
    }
}
