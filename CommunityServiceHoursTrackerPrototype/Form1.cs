using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Word;
using System.Configuration;

namespace CommunityServiceHoursTracker
{
    public partial class Form1 : Form
    {
        private string connStr = ConfigurationManager.ConnectionStrings["DB"].ConnectionString;
        private MySqlConnection thisConnection;
        bool isUpdate = true;
        bool resettingTime = false;

        //strings to hold the hours and times in the time amount
        string hour = string.Empty;
        string min = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            //Fill controls

            //Retrieve list of volunteers
            //MySqlConnection thisConnection = new MySqlConnection(connStr);
            FillPerson1DDL();
            FillPerson2DDL();
            FillPerson3DDL();

            //Retrieve list of cases for selected volunteer
            if(SelectVolunteer1DDL.Items.Count > 0){
                FillCases1DDL(SelectVolunteer1DDL.SelectedValue.ToString());
                FillCases2DDL(SelectVolunteer2DDL.SelectedValue.ToString());
                FillCases3DDL(SelectVolunteer3DDL.SelectedValue.ToString());

                //Fill current volunteer information
                FillVolunteerInfoFields();

                //Fill current case information
                FillCaseInfoFields();
            }

            CheckEnterTimeTabEnabling();
            CheckEnabling(); 
            CheckReportsTabEnabling();
        }

        private void FillPerson1DDL()
        {
            thisConnection = new MySqlConnection(connStr);
            try
            {
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT Concat(LName, ', ', FName, ' ', MInitial) AS VolunteerName, PersonID FROM person WHERE Status = True ORDER BY LName;";

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataTable table = new DataTable();
                dataAdapter.Fill(table);

                SelectVolunteer1DDL.Text = "";
                SelectVolunteer1DDL.DataSource = table;
                SelectVolunteer1DDL.DisplayMember = "VolunteerName";
                SelectVolunteer1DDL.ValueMember = "PersonID";
                if (table.Rows.Count > 0)
                {
                    SelectVolunteer1DDL.SelectedIndex = 0;
                }
                else
                {
                    SelectCase1DDL.DataSource = null;
                    CheckEnterTimeTabEnabling();
                }
                ResetEnterTime();
            }
            catch (MySqlException ee)
            {
                Console.WriteLine(ee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        private void FillPerson2DDL()
        {
            thisConnection = new MySqlConnection(connStr);
            try
            {
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT Concat(LName, ', ', FName, ' ', MInitial) AS VolunteerName, PersonID FROM person ORDER BY LName;";

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataTable table = new DataTable();
                dataAdapter.Fill(table);

                SelectVolunteer2DDL.Text = "";
                SelectVolunteer2DDL.DataSource = table;
                SelectVolunteer2DDL.DisplayMember = "VolunteerName";
                SelectVolunteer2DDL.ValueMember = "PersonID";

                if (table.Rows.Count > 0)
                {
                    SelectVolunteer2DDL.SelectedIndex = 0;
                }
                else
                {
                    SelectCase2DDL.DataSource = null;
                    CheckEnabling();
                }
                ResetEnterTime();
            }
            catch (MySqlException ee)
            {
                Console.WriteLine(ee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        private void FillPerson3DDL()
        {
            thisConnection = new MySqlConnection(connStr);
            try
            {
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT Concat(LName, ', ', FName, ' ', MInitial) AS VolunteerName, PersonID FROM person WHERE Status = True ORDER BY LName;";

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataTable table = new DataTable();
                dataAdapter.Fill(table);

                SelectVolunteer3DDL.Text = "";
                SelectVolunteer3DDL.DataSource = table;
                SelectVolunteer3DDL.DisplayMember = "VolunteerName";
                SelectVolunteer3DDL.ValueMember = "PersonID";

                if (table.Rows.Count > 0)
                {
                    SelectVolunteer3DDL.SelectedIndex = 0;
                }
                else
                {
                    SelectCase3DDL.DataSource = null;
                    CheckReportsTabEnabling();
                }
                ResetEnterTime();
            }
            catch (MySqlException ee)
            {
                Console.WriteLine(ee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        private void FillCases1DDL(string selectedPerson)
        {
            SelectCase1DDL.Text = "";
            SelectCase1DDL.DataBindings.Clear();
            try
            {
                thisConnection = new MySqlConnection(connStr);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                
                thisCommand.CommandText = "SELECT CaseNum, CaseID FROM cases WHERE Cases.PersonID = " + selectedPerson + " AND Status = True ORDER BY CaseNum;";

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataTable table = new DataTable();
                dataAdapter.Fill(table);


                if (table.Rows.Count > 0)
                {
                    SelectCase1DDL.Text = table.Rows[0].ItemArray[0].ToString();
                    //SelectCase1DDL.SelectedIndex = 0;
                }

                SelectCase1DDL.Text = "";
                SelectCase1DDL.DataSource = table;
                SelectCase1DDL.DisplayMember = "CaseNum";
                SelectCase1DDL.ValueMember = "CaseID";
                if (SelectCase1DDL.Items.Count > 0)
                {
                    SelectCase1DDL.SelectedIndex = 0;
                }

                CheckEnterTimeTabEnabling();
                ResetEnterTime();
            }
            catch (MySqlException ee)
            {
                Console.WriteLine(ee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        private void FillCases2DDL(string selectedPerson)
        {
            SelectCase2DDL.Text = "";
            SelectCase2DDL.DataSource = null;
            try
            {
                thisConnection = new MySqlConnection(connStr);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();

                thisCommand.CommandText = "SELECT CaseNum, CaseID FROM cases WHERE Cases.PersonID = " + selectedPerson + " ORDER BY CaseNum;";

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataTable table = new DataTable();
                dataAdapter.Fill(table);


                if (table.Rows.Count > 0)
                {
                    SelectCase2DDL.Text = table.Rows[0].ItemArray[0].ToString();
                    //SelectCase2DDL.SelectedIndex = 0;
                }

                SelectCase2DDL.Text = "";
                SelectCase2DDL.DataSource = table;
                SelectCase2DDL.DisplayMember = "CaseNum";
                SelectCase2DDL.ValueMember = "CaseID";

                CheckEnabling();
                ResetEnterTime();
            }
            catch (MySqlException ee)
            {
                Console.WriteLine(ee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        private void FillCases3DDL(string selectedPerson)
        {
            SelectCase3DDL.Text = "";
            SelectCase3DDL.DataSource = null;
            try
            {
                thisConnection = new MySqlConnection(connStr);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();

                thisCommand.CommandText = "SELECT CaseNum, CaseID FROM cases WHERE Cases.PersonID = " + selectedPerson + " AND Status = True ORDER BY CaseNum;";

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataTable table = new DataTable();
                dataAdapter.Fill(table);


                if (table.Rows.Count > 0)
                {
                    SelectCase3DDL.Text = table.Rows[0].ItemArray[0].ToString();
                    //SelectCase3DDL.SelectedIndex = 0;
                }

                SelectCase3DDL.Text = "";
                SelectCase3DDL.DataSource = table;
                SelectCase3DDL.DisplayMember = "CaseNum";
                SelectCase3DDL.ValueMember = "CaseID";

                CheckReportsTabEnabling();
                ResetEnterTime();
            }
            catch (MySqlException ee)
            {
                Console.WriteLine(ee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        //Enter Time Tab:
        private void CheckEnterTimeTabEnabling()
        {
            if (SelectCase1DDL.Items.Count.Equals(0))
            {
                inMonth.Enabled = false;
                inDay.Enabled = false;
                inYear.Enabled = false;
                inHour.Enabled = false;
                inMin.Enabled = false;
                inAmPm.Enabled = false;

                outMonth.Enabled = false;
                outDay.Enabled = false;
                outYear.Enabled = false;
                outHour.Enabled = false;
                outMin.Enabled = false;
                outAmPm.Enabled = false;

                TotalHoursTextBox.Enabled = false;
                SaveTimeButton.Enabled = false;

                //MessageBox.Show("This tab is disabled because the selected volunteer does not have an active case assigned");
            }
            else if (SelectCase1DDL.Items.Count > 0)
            {
                inMonth.Enabled = true;
                inDay.Enabled = true;
                inYear.Enabled = true;
                inHour.Enabled = true;
                inMin.Enabled = true;
                inAmPm.Enabled = true;

                outMonth.Enabled = true;
                outDay.Enabled = true;
                outYear.Enabled = true;
                outHour.Enabled = true;
                outMin.Enabled = true;
                outAmPm.Enabled = true;

                SaveTimeButton.Enabled = true;
            }
        }

        private void SelectVolunteer1DDL_SelectedIndexChanged(object sender, EventArgs e)
        {
            string currPerson = SelectVolunteer1DDL.SelectedValue.ToString();
            txtTotalHours.Text = "00:00";
            txtNeeded.Text = "00:00";
            grdViewHours.DataBindings.Clear();
            grdViewHours.DataSource = null;
            FillCases1DDL(currPerson);
            //FillVolunteerInfoFields();
            CheckEnterTimeTabEnabling();
           
        }

        private void dateValueChanged(object sender, EventArgs e)
        {
            if (resettingTime)
            {
                return; //we are resetting the time programmatically so we don't want to compare...only when done by a person
            }
            try
            {
                if (inMin.Text.Length == 1)
                {
                    inMin.Text = "0" + inMin.Text;
                }
                if (outMin.Text.Length == 1)
                {
                    outMin.Text = "0" + outMin.Text;
                }
                TimeSpan ts = getTimeOut().Subtract(getTimeIn());

                string hours = Convert.ToInt32(ts.Hours).ToString();
                string mins = Convert.ToInt32(ts.Minutes).ToString();
                hour = hours;
                min = mins;
                
                if ((Convert.ToInt32(mins) < 10) && (Convert.ToInt32(mins) >= 0))
                {
                    mins = "0" + mins;
                }
                if ((Convert.ToInt32(hours) < 10) && (Convert.ToInt32(hours) >= 0))
                {
                    hours = "0" + hours;
                }
                //only update time if it is a correct time

                if (Convert.ToInt32(hours) >= 0 && Convert.ToInt32(mins) >= 0)
                {
                    TotalHoursTextBox.Text = hours + ":" + mins;
                }
                else
                {
                    TotalHoursTextBox.Text = "00:00";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool equalDateTimes()
        {
            bool check = false;
            if (getTimeIn().Date.Equals(getTimeOut().Date))
            {
                if (getTimeIn().Hour.Equals(getTimeOut().Hour))
                {
                    if (getTimeIn().Minute.Equals(getTimeOut().Minute))
                    {
                        check = true;
                    }
                }
            }
            return check;
        }

        private int CheckDateTimePickers()
        {
            int pass = 0;

            if (tabControl1.SelectedTab.Equals(tabPage1))
            {
                if ((DateTime.Compare(getTimeIn(), getTimeOut()) == 0)
                    || (DateTime.Compare(getTimeIn(), getTimeOut()) < 0)
                    || equalDateTimes())
                {
                    TotalHoursTextBox.Text = hour + ":" + min;
                    pass = 1;
                }
                else if (DateTime.Compare(getTimeIn(), getTimeOut()) > 0)
                {
                    MessageBox.Show("Time in is not before time out.");

                    TotalHoursTextBox.Text = "00:00";
                }
            }
            return pass;
        }
        
        private void SaveTimeButton_Click(object sender, EventArgs e)
        {


            int pass = CheckDateTimePickers();
            
            try
            {
                if (pass == 1)
                {
                    thisConnection = new MySqlConnection(connStr);
                    thisConnection.Open();
                    MySqlCommand thisCommand = thisConnection.CreateCommand();
                    thisCommand.CommandText = "INSERT INTO event (TimeIn, TimeOut, CaseID) VALUES(@StartTime, @EndTime, @CaseID);";
                    thisCommand.Parameters.Add("@StartTime", MySqlDbType.DateTime);
                    thisCommand.Parameters["@StartTime"].Value = getTimeIn();
                    thisCommand.Parameters.Add("@EndTime", MySqlDbType.DateTime);
                    thisCommand.Parameters["@EndTime"].Value = getTimeOut();
                    thisCommand.Parameters.Add("@CaseID", MySqlDbType.Int32);
                    thisCommand.Parameters["@CaseID"].Value = SelectCase1DDL.SelectedValue;
                    thisCommand.Prepare();
                    thisCommand.ExecuteNonQuery();
                    MessageBox.Show("Your event was saved successfully.");
                    fillTimeGrid();
                    ResetEnterTime();
                }
            }
            catch (MySqlException ee)
            {
                MessageBox.Show("An error occured connecting to the database!");
            }
            catch (Exception eee)
            {
                MessageBox.Show(eee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        //Volunteers Tab:
        private void SelectVolunteer2DDL_SelectedIndexChanged(object sender, EventArgs e)
        {
            string currPerson = SelectVolunteer2DDL.SelectedValue.ToString();
            FillCases2DDL(currPerson);
            FillVolunteerInfoFields();
            CheckEnabling();
        }

        private void AddNewVolunteerButton_Click(object sender, EventArgs e)
        {
            if (AddNewVolunteerButton.Text.Equals("Add New Volunteer"))
            {
                isUpdate = false;
                SelectVolunteer2DDL.Enabled = false;
                AddNewVolunteerButton.Text = "Cancel Add New Volunteer";
                setAddNewVolunteerButtons(false); 
                txtFirstName.Text = "";
                txtMiddleInitial.Text = "";
                txtLastName.Text = "";
                txtAddress.Text = "";
                lastContactDate.Value = DateTime.Now;
                checkBoxActiveVolunteer.Checked = false;
                
            }
            else if (AddNewVolunteerButton.Text.Equals("Cancel Add New Volunteer"))
            {
                isUpdate = true;
                SelectVolunteer2DDL.Enabled = true;
                AddNewVolunteerButton.Text = "Add New Volunteer";
                setAddNewVolunteerButtons(true);
                FillVolunteerInfoFields();
            }
        }

        private void FillVolunteerInfoFields()
        {
            //Updating volunteer info
            try
            {
                thisConnection = new MySqlConnection(connStr);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT FName, MInitial, LName, Address, LastContactDay, Status FROM person WHERE PersonID = '" + SelectVolunteer2DDL.SelectedValue + "';";

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataTable table = new DataTable();
                dataAdapter.Fill(table);

                txtFirstName.Text = table.Rows[0].ItemArray[0].ToString();
                txtMiddleInitial.Text = table.Rows[0].ItemArray[1].ToString();
                txtLastName.Text = table.Rows[0].ItemArray[2].ToString();
                txtAddress.Text = table.Rows[0].ItemArray[3].ToString();
                lastContactDate.Value = Convert.ToDateTime(table.Rows[0].ItemArray[4]);
                if ((table.Rows[0].ItemArray[5].ToString()).Equals("0"))
                {
                    checkBoxActiveVolunteer.Checked = false;
                }
                else if ((table.Rows[0].ItemArray[5].ToString()).Equals("1"))
                {
                    checkBoxActiveVolunteer.Checked = true;
                }
            }
            catch (MySqlException ee)
            {
                Console.WriteLine(ee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        private void SaveVolunteerButton_Click(object sender, EventArgs e)
        {
            string volunteerName = txtLastName.Text + ", " + txtFirstName.Text + " " + txtMiddleInitial.Text;
            if (AddNewVolunteerButton.Text.Equals("Add New Volunteer"))
            {
                //UPDATE
                try
                {
                    thisConnection = new MySqlConnection(connStr);
                    thisConnection.Open();
                    MySqlCommand thisCommand = thisConnection.CreateCommand();
                    bool volStatus = false;
                    if (checkBoxActiveVolunteer.Checked)
                    {
                        volStatus = true;
                    }
                    //thisCommand.CommandText = "UPDATE person SET FName='" + txtFirstName.Text + "', MInitial='" + txtMiddleInitial.Text + "', LName='" + txtLastName.Text + "', Address='" + txtAddress.Text + "', LastContactDay=str_to_date('" + lastContactDate.Value + "', '%m/%e/%Y %h:%i:%s %p'), Status=" + volStatus + " WHERE PersonID = " + SelectVolunteer2DDL.SelectedValue + ";";
                    thisCommand.CommandText = "UPDATE person SET FName= @FirstName , MInitial= @MiddleInit , LName= @LastName, Address= @Addr, LastContactDay= @ContactDate , Status= @Status WHERE PersonID = @PersonID;";
                    thisCommand.Parameters.Add("@FirstName", MySqlDbType.VarChar, txtFirstName.Text.Length);
                    thisCommand.Parameters["@FirstName"].Value = txtFirstName.Text;
                    thisCommand.Parameters.Add("@MiddleInit", MySqlDbType.VarChar, txtMiddleInitial.Text.Length);
                    thisCommand.Parameters["@MiddleInit"].Value = txtMiddleInitial.Text;
                    thisCommand.Parameters.Add("@LastName", MySqlDbType.VarChar, txtLastName.Text.Length);
                    thisCommand.Parameters["@LastName"].Value = txtLastName.Text;
                    thisCommand.Parameters.Add("@Addr", MySqlDbType.VarChar, txtAddress.Text.Length);
                    thisCommand.Parameters["@Addr"].Value = txtAddress.Text;
                    thisCommand.Parameters.Add("@ContactDate", MySqlDbType.DateTime);
                    thisCommand.Parameters["@ContactDate"].Value = lastContactDate.Value;
                    thisCommand.Parameters.Add("@Status", MySqlDbType.Bit);
                    thisCommand.Parameters["@Status"].Value = volStatus;
                    thisCommand.Parameters.Add("@PersonID", MySqlDbType.Int32);
                    thisCommand.Parameters["@PersonID"].Value = SelectVolunteer2DDL.SelectedValue;
                    thisCommand.Prepare();
                    thisCommand.ExecuteNonQuery();
                    
                    //refresh drop down list data
                    int enterId = -1;
                    int volunteerId = -1;
                    int reportId = -1;
                    if (SelectVolunteer1DDL.FindString(SelectVolunteer1DDL.Text) > 0)
                    {
                        enterId = Convert.ToInt32(SelectVolunteer1DDL.SelectedValue);
                    }
                    if (SelectVolunteer2DDL.FindString(SelectVolunteer2DDL.Text) > 0)
                    {
                        volunteerId = Convert.ToInt32(SelectVolunteer2DDL.SelectedValue);
                    }
                    if (SelectVolunteer3DDL.FindString(SelectVolunteer3DDL.Text) > 0)
                    {
                        reportId = Convert.ToInt32(SelectVolunteer3DDL.SelectedValue);
                    }
                   
                    FillPerson1DDL();
                    if (enterId >= 0)
                    {
                        SelectVolunteer1DDL.SelectedValue = enterId;
                    }

                    FillPerson2DDL();
                    if (volunteerId > 0)
                    {
                        SelectVolunteer2DDL.SelectedValue = volunteerId;
                    }
                    FillPerson3DDL();
                    if (reportId >= 0)
                    {
                        SelectVolunteer3DDL.SelectedValue = reportId;
                    }


                    MessageBox.Show("Volunteer " + volunteerName + " was updated successfully.");
                }
                catch (MySqlException ee)
                {
                    MessageBox.Show("An error occured connecting to the database!");
                }
                catch (Exception eee)
                {
                    MessageBox.Show(eee.Message);
                }
                finally
                {
                    thisConnection.Close();
                }
            }
            else if (AddNewVolunteerButton.Text.Equals("Cancel Add New Volunteer"))
            {
                //INSERT
                try
                {
                    thisConnection = new MySqlConnection(connStr);
                    thisConnection.Open();
                    MySqlCommand thisCommand = thisConnection.CreateCommand();

                    bool volStatus = false;
                    if (checkBoxActiveVolunteer.Checked)
                    {
                        volStatus = true;
                    }
                    //thisCommand.CommandText = "INSERT INTO person (FName, MInitial, LName, Address, LastContactDay, Status) VALUES('" + txtFirstName.Text + "', '" + txtMiddleInitial.Text + "', '" + txtLastName.Text + "', '" + txtAddress.Text + "',  str_to_date('" + lastContactDate.Value + "', '%m/%e/%Y %h:%i:%s %p'), " + volStatus + ");";
                    thisCommand.CommandText = "INSERT INTO person (FName, MInitial, LName, Address, LastContactDay, Status) VALUES(@FirstName, @MiddleInit, @LastName, @Addr, @ContactDate, @Status);";
                    thisCommand.Parameters.Add("@FirstName", MySqlDbType.VarChar, txtFirstName.Text.Length);
                    thisCommand.Parameters["@FirstName"].Value = txtFirstName.Text;
                    thisCommand.Parameters.Add("@MiddleInit", MySqlDbType.VarChar, txtMiddleInitial.Text.Length);
                    thisCommand.Parameters["@MiddleInit"].Value = txtMiddleInitial.Text;
                    thisCommand.Parameters.Add("@LastName", MySqlDbType.VarChar, txtLastName.Text.Length);
                    thisCommand.Parameters["@LastName"].Value = txtLastName.Text;
                    thisCommand.Parameters.Add("@Addr", MySqlDbType.VarChar, txtAddress.Text.Length);
                    thisCommand.Parameters["@Addr"].Value = txtAddress.Text;
                    thisCommand.Parameters.Add("@ContactDate", MySqlDbType.DateTime);
                    thisCommand.Parameters["@ContactDate"].Value = lastContactDate.Value;
                    thisCommand.Parameters.Add("@Status", MySqlDbType.Bit);
                    thisCommand.Parameters["@Status"].Value = volStatus;
                    thisCommand.Prepare();
                    thisCommand.ExecuteNonQuery();

                    //refresh drop down list data
                    int enterId = Convert.ToInt32(SelectVolunteer1DDL.SelectedValue);
                    int reportId = Convert.ToInt32(SelectVolunteer3DDL.SelectedValue);
                    bool refresh = true;
                    if (SelectVolunteer1DDL.Items.Count == 0)
                    {
                        refresh = false;
                    }
                    FillPerson1DDL();
                    if (refresh)
                    {
                        SelectVolunteer1DDL.SelectedValue = enterId;
                        FillCases1DDL(enterId.ToString());
                    }
                         
                    FillPerson2DDL();
                    
                    FillPerson3DDL();
                    if (refresh)
                    {
                        SelectVolunteer3DDL.SelectedValue = reportId;
                        FillCases3DDL(reportId.ToString());
                    }

                    AddNewVolunteerButton.Text = "Add New Volunteer";
                    setAddNewVolunteerButtons(true);
                    SelectVolunteer2DDL.Enabled = true;
                    MessageBox.Show("Volunteer " + volunteerName + " was added successfully.");
                }
                catch (MySqlException ee)
                {
                    MessageBox.Show("An error occured connecting to the database!");
                }
                catch (Exception eee)
                {
                    MessageBox.Show(eee.Message);
                }
                finally
                {
                    thisConnection.Close();
                }

                string volunteerID = "";
                try
                {
                    thisConnection = new MySqlConnection(connStr);
                    thisConnection.Open();
                    MySqlCommand thisCommand = thisConnection.CreateCommand();
                    thisCommand.CommandText = "SELECT MAX(PersonID) FROM person;";

                    MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                    MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                    DataTable table = new DataTable();
                    dataAdapter.Fill(table);

                    volunteerID = table.Rows[0].ItemArray[0].ToString();
                }
                catch (MySqlException ee)
                {
                    MessageBox.Show("An error occured connecting to the database!");
                }
                catch (Exception eee)
                {
                    MessageBox.Show(eee.Message);
                }
                finally
                {
                 
                    thisConnection.Close();
                }
                
                SelectVolunteer2DDL.SelectedValue = volunteerID;
                FillVolunteerInfoFields(); 
                FillCases2DDL(volunteerID);
                FillCaseInfoFields();
            }
        }

        private void FillCaseInfoFields()
        {
            try
            {
                thisConnection = new MySqlConnection(connStr);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT CaseNum, HoursAssigned, AgencyContactDate, EstimatedCompletionDate, Status, WeeklyReq FROM cases WHERE PersonID = '" + SelectVolunteer2DDL.SelectedValue + "' AND CaseID = '" + SelectCase2DDL.SelectedValue + "';";

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataTable table = new DataTable();
                dataAdapter.Fill(table);

                if (table.Rows.Count > 0)
                {
                    txtCaseNum.Text = table.Rows[0].ItemArray[0].ToString();
                    txtHoursAssigned.Text = table.Rows[0].ItemArray[1].ToString();
                    agencyContactDate.Value = Convert.ToDateTime(table.Rows[0].ItemArray[2]);
                    estimatedCompletionDate.Value = Convert.ToDateTime(table.Rows[0].ItemArray[3]);
                    txtWeeklyReq.Text = Convert.ToString(table.Rows[0].ItemArray[5]);

                    if ((table.Rows[0].ItemArray[4].ToString()).Equals("0"))
                    {
                        checkBoxActiveCase.Checked = false;
                    }
                    else if ((table.Rows[0].ItemArray[4].ToString()).Equals("1"))
                    {
                        checkBoxActiveCase.Checked = true;
                    }
                }
                else
                {
                    txtCaseNum.Text = "";
                    txtHoursAssigned.Text = "";
                    txtWeeklyReq.Text = "";
                    agencyContactDate.Value = DateTime.Now;
                    estimatedCompletionDate.Value = DateTime.Now;
                    checkBoxActiveCase.Checked = false;
                }
            }
            catch (MySqlException ee)
            {
                Console.WriteLine(ee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        private void SelectCase2DDL_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillCaseInfoFields();
        }

        private void EditCaseButton_Click(object sender, EventArgs e)
        {
            if (EditCaseButton.Text.Equals("Edit Selected Case"))
            {
                enableCase(true);
                setEditCaseButtons(false);
                SaveNewCaseButton.Text = "Save Updated Case";
                EditCaseButton.Text = "Cancel Editing Case";
            }
            else if (EditCaseButton.Text.Equals("Cancel Editing Case"))
            {
                enableCase(false);
                setEditCaseButtons(true);
                SaveNewCaseButton.Text = "Save";
                EditCaseButton.Text = "Edit Selected Case";
            }
        }

        private void CloseCaseButton_Click(object sender, EventArgs e)
        {
            try
            {
                string theCaseNum = SelectCase2DDL.SelectedText.ToString();
                thisConnection = new MySqlConnection(connStr);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                bool theStatus = false;

                thisCommand.CommandText = "UPDATE cases SET Status= @Status WHERE CaseID = @CaseID;";
                thisCommand.Parameters.Add("@Status", MySqlDbType.Bit);
                thisCommand.Parameters["@Status"].Value = theStatus;
                thisCommand.Parameters.Add("@CaseID", MySqlDbType.Int32);
                thisCommand.Parameters["@CaseID"].Value = SelectCase2DDL.SelectedValue;
                thisCommand.Prepare();
                thisCommand.ExecuteNonQuery();
                MessageBox.Show("Case number " + theCaseNum + " was closed successfully.");
                FillCaseInfoFields();
                FillCases1DDL(SelectVolunteer1DDL.SelectedValue.ToString());
                FillCases3DDL(SelectVolunteer3DDL.SelectedValue.ToString());
            }
            catch (MySqlException ee)
            {
                MessageBox.Show("An error occured connecting to the database!");
            }
            catch (Exception eee)
            {
                MessageBox.Show(eee.Message);
            }
            finally
            {
                thisConnection.Close();
            }
        }

        private void SaveNewCaseButton_Click(object sender, EventArgs e)
        {
            enableCase(false);
            setAddNewCaseButtons(true);
            btnAddCase.Enabled = true;
            string volunteerCase = txtCaseNum.Text;
            bool caseStatus = false;
            if (checkBoxActiveCase.Checked)
            {
                caseStatus = true;
            }

            string enterCaseId = Convert.ToString(SelectCase1DDL.SelectedValue);
            string volunteerCaseId = Convert.ToString(SelectCase2DDL.SelectedValue);
            string reportCaseId = Convert.ToString(SelectCase3DDL.SelectedValue);

            if (EditCaseButton.Text.Equals("Cancel Editing Case"))
            {
                //Update
                try
                {
                    thisConnection = new MySqlConnection(connStr);
                    thisConnection.Open();
                    MySqlCommand thisCommand = thisConnection.CreateCommand();
                    thisCommand.CommandText = "UPDATE cases SET CaseNum= @CaseNum , HoursAssigned= @HoursAssigned, AgencyContactDate= @ContactDate, EstimatedCompletionDate= @EstCompletionDate, Status= @CaseStatus, WeeklyReq = @WeeklyReq WHERE CaseID = @CaseID;";
                    thisCommand.Parameters.Add("@CaseNum", MySqlDbType.VarChar, txtCaseNum.Text.Length);
                    thisCommand.Parameters["@CaseNum"].Value = txtCaseNum.Text;
                    thisCommand.Parameters.Add("@HoursAssigned", MySqlDbType.Decimal);
                    thisCommand.Parameters["@HoursAssigned"].Value = Convert.ToDecimal(txtHoursAssigned.Text);
                    thisCommand.Parameters.Add("@ContactDate", MySqlDbType.DateTime);
                    thisCommand.Parameters["@ContactDate"].Value = agencyContactDate.Value;
                    thisCommand.Parameters.Add("@EstCompletionDate", MySqlDbType.DateTime);
                    thisCommand.Parameters["@EstCompletionDate"].Value = estimatedCompletionDate.Value;
                    thisCommand.Parameters.Add("@CaseStatus", MySqlDbType.Bit);
                    thisCommand.Parameters["@CaseStatus"].Value = caseStatus;
                    thisCommand.Parameters.Add("@WeeklyReq", MySqlDbType.Decimal);
                    thisCommand.Parameters["@WeeklyReq"].Value = Convert.ToDecimal(txtWeeklyReq.Text);
                    thisCommand.Parameters.Add("@CaseID", MySqlDbType.Int32);
                    thisCommand.Parameters["@CaseID"].Value = SelectCase2DDL.SelectedValue;
                    thisCommand.Prepare();
                    thisCommand.ExecuteNonQuery();
                    //FillPerson2DDL();
                   
                    MessageBox.Show("Case number " + volunteerCase + " was updated successfully.");
                    EditCaseButton.Text = "Edit Selected Case";
                }
                catch (MySqlException ee)
                {
                    MessageBox.Show("An error occured connecting to the database!");
                }
                catch (Exception eee)
                {
                    MessageBox.Show(eee.Message);
                }
                finally
                {
                    thisConnection.Close();
                }

                //refresh drop down list data    
                int enterId = Convert.ToInt32(SelectVolunteer1DDL.SelectedValue);
                int volunteerId = Convert.ToInt32(SelectVolunteer2DDL.SelectedValue);
                int reportId = Convert.ToInt32(SelectVolunteer3DDL.SelectedValue);

                //Refresh select cases/case fields
                FillCases1DDL(enterId.ToString());
                if (!enterCaseId.Equals(""))
                {
                    SelectCase1DDL.SelectedValue = enterCaseId;
                }

                FillCases2DDL(volunteerId.ToString());
                if (!volunteerCase.Equals(""))
                {
                    SelectCase2DDL.SelectedValue = volunteerCaseId;
                    FillCaseInfoFields();
                }

                FillCases3DDL(reportId.ToString());
                if (!reportCaseId.Equals(""))
                {
                    SelectCase3DDL.SelectedValue = reportCaseId;
                }
            }
            else if(btnAddCase.Text.Equals("Cancel Add Case"))
            {
                //Insert
                try
                {
                    thisConnection = new MySqlConnection(connStr);
                    thisConnection.Open();
                    MySqlCommand thisCommand = thisConnection.CreateCommand();
                    thisCommand.CommandText = "INSERT INTO cases (CaseNum, HoursAssigned, AgencyContactDate, EstimatedCompletionDate, Status, PersonID, WeeklyReq) VALUES(@CaseNum, @HoursAssigned, @AgencyContactDate, @EstCompletionDate, @CaseStatus, @PersonID, @WeeklyReq);";
                    thisCommand.Parameters.Add("@CaseNum", MySqlDbType.VarChar, txtCaseNum.Text.Length);
                    thisCommand.Parameters["@CaseNum"].Value = txtCaseNum.Text;
                    thisCommand.Parameters.Add("@HoursAssigned", MySqlDbType.Decimal);
                    thisCommand.Parameters["@HoursAssigned"].Value = txtHoursAssigned.Text;
                    thisCommand.Parameters.Add("@AgencyContactDate", MySqlDbType.DateTime);
                    thisCommand.Parameters["@AgencyContactDate"].Value = agencyContactDate.Value;
                    thisCommand.Parameters.Add("@EstCompletionDate", MySqlDbType.DateTime);
                    thisCommand.Parameters["@EstCompletionDate"].Value = estimatedCompletionDate.Value;
                    thisCommand.Parameters.Add("@CaseStatus", MySqlDbType.Bit);
                    thisCommand.Parameters["@CaseStatus"].Value = caseStatus;
                    thisCommand.Parameters.Add("@PersonID", MySqlDbType.Int32);
                    thisCommand.Parameters["@PersonID"].Value = SelectVolunteer2DDL.SelectedValue;
                    thisCommand.Parameters.Add("@WeeklyReq", MySqlDbType.Decimal);
                    thisCommand.Parameters["@WeeklyReq"].Value = Convert.ToDecimal(txtWeeklyReq.Text);
                    thisCommand.Prepare();
                    thisCommand.ExecuteNonQuery();
                    MessageBox.Show("Your event was saved successfully.");
                    btnAddCase.Text = "Add New Case";
                }
                catch (MySqlException ee)
                {
                    MessageBox.Show("An error occured connecting to the database!");
                }
                catch (Exception eee)
                {
                    MessageBox.Show(eee.Message);
                }
                finally
                {
                    thisConnection.Close();
                }

                //string caseID = txtCaseNum.Text;
                string caseID = "";
                try
                {
                    thisConnection = new MySqlConnection(connStr);
                    thisConnection.Open();
                    MySqlCommand thisCommand = thisConnection.CreateCommand();
                    thisCommand.CommandText = "SELECT MAX(CaseID) FROM cases;";

                    MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                    MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                    DataTable table = new DataTable();
                    dataAdapter.Fill(table);

                    caseID = table.Rows[0].ItemArray[0].ToString();
                }
                catch (MySqlException ee)
                {
                    MessageBox.Show("An error occured connecting to the database!");
                }
                catch (Exception eee)
                {
                    MessageBox.Show(eee.Message);
                }
                finally
                {
                    thisConnection.Close();
                }

                //refresh drop down list data    
                int enterId = Convert.ToInt32(SelectVolunteer1DDL.SelectedValue);
                int volunteerId = Convert.ToInt32(SelectVolunteer2DDL.SelectedValue);
                int reportId = Convert.ToInt32(SelectVolunteer3DDL.SelectedValue);

                //Refresh select cases/case fields
                FillCases1DDL(enterId.ToString());
                if (!enterCaseId.Equals(""))
                {
                    SelectCase1DDL.SelectedValue = enterCaseId;
                }
               

                FillCases2DDL(volunteerId.ToString());
                SelectCase2DDL.SelectedValue = caseID;
                FillCaseInfoFields();

                FillCases3DDL(reportId.ToString());

                if (!reportCaseId.Equals(""))
                {
                    SelectCase3DDL.SelectedValue = reportCaseId;
                }
                
            }
        }

        private void btnAddCase_Click(object sender, EventArgs e)
        {
            if (btnAddCase.Text.Equals("Add New Case"))
            {
                enableCase(true);
                setAddNewCaseButtons(false);
                txtCaseNum.Text = "";
                txtHoursAssigned.Text = "";
                txtWeeklyReq.Text = "";
                agencyContactDate.Value = DateTime.Now;
                estimatedCompletionDate.Value = DateTime.Now;
                checkBoxActiveCase.Checked = true;
                btnAddCase.Text = "Cancel Add Case";
            }
            else if (btnAddCase.Text.Equals("Cancel Add Case"))
            {
                enableCase(false);
                setAddNewCaseButtons(true);
                btnAddCase.Text = "Add New Case";
                FillCaseInfoFields();
            }
        }

        private void setAddNewVolunteerButtons(bool enable)
        {
            SelectCase2DDL.Enabled = enable;
            EditCaseButton.Enabled = enable;
            btnAddCase.Enabled = enable;
            CloseCaseButton.Enabled = enable;
        }

        private void enableCase(bool enable)
        {
            txtCaseNum.Enabled = enable;
            txtHoursAssigned.Enabled = enable;
            agencyContactDate.Enabled = enable;
            estimatedCompletionDate.Enabled = enable;
            checkBoxActiveCase.Enabled = enable;
            SaveNewCaseButton.Enabled = enable;
            txtWeeklyReq.Enabled = enable;
        }

        private void setEditCaseButtons(bool enable)
        {
            SelectCase2DDL.Enabled = enable;
            btnAddCase.Enabled = enable;
            CloseCaseButton.Enabled = enable;
            SelectVolunteer2DDL.Enabled = enable;
            AddNewVolunteerButton.Enabled = enable;
            SaveVolunteerButton.Enabled = enable;
        }

        private void setAddNewCaseButtons(bool enable)
        {
            EditCaseButton.Enabled = enable;
            SelectCase2DDL.Enabled = enable;
            CloseCaseButton.Enabled = enable;
            SelectVolunteer2DDL.Enabled = enable;
            AddNewVolunteerButton.Enabled = enable;
            SaveVolunteerButton.Enabled = enable;
        }

        private void CheckEnabling()
        {
            if (SelectCase2DDL.Items.Count.Equals(0))
            {
                EditCaseButton.Enabled = false;
                CloseCaseButton.Enabled = false;
            }
            else if (SelectCase2DDL.Items.Count > 0)
            {
                EditCaseButton.Enabled = true;
                CloseCaseButton.Enabled = true;
            }
        }
        
        //Reports Tab:
        private void SelectVolunteer3DDL_SelectedIndexChanged(object sender, EventArgs e)
        {
            string currPerson = SelectVolunteer3DDL.SelectedValue.ToString();
            FillCases3DDL(currPerson);
            FillVolunteerInfoFields();
            CheckReportsTabEnabling();
        }

        private void CheckReportsTabEnabling()
        {
            if (SelectCase3DDL.Items.Count.Equals(0))
            {
                dateTimePicker6.Enabled = false;
                CreateMonthlyReportButton.Enabled = false;
                CreateTimeSheetReportButton.Enabled = false;
                //MessageBox.Show("This tab is disabled because the selected volunteer does not have an active case assigned");
            }
            else if (SelectCase3DDL.Items.Count > 0)
            {
                dateTimePicker6.Enabled = true;
                CreateMonthlyReportButton.Enabled = true;
                CreateTimeSheetReportButton.Enabled = true;
            }
        }

        private void CreateMonthlyReportButton_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.ApplicationClass app = new Microsoft.Office.Interop.Word.ApplicationClass();
                Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();

                object missing = System.Reflection.Missing.Value;
                object b = false;
                object c = 0;
                object d = false;
                doc = app.Documents.Add(ref missing, ref b, ref c, ref d);
                doc.Activate();


                string connectionString = ConfigurationManager.ConnectionStrings["DB"].ConnectionString;
                thisConnection = new MySqlConnection(connectionString);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = thisCommand.CommandText = "SELECT CaseNum, HoursAssigned, AgencyContactDate, EstimatedCompletionDate, Status FROM cases WHERE PersonID = '" + SelectVolunteer3DDL.SelectedValue + "' AND CaseID = '" + SelectCase3DDL.SelectedValue + "';";
                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataSet DS = new DataSet();
                dataAdapter.Fill(DS, "CaseInfo");

                thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT FName, MInitial, LName, Address, LastContactDay, Status FROM person WHERE PersonID = '" + SelectVolunteer3DDL.SelectedValue + "';";
                dataAdapter.SelectCommand = thisCommand;
                dataAdapter.Fill(DS, "PersonInfo");

                //for monthly report
                DateTime combo = dateTimePicker6.Value;
                DateTime beginMonth = new DateTime(combo.Year, combo.Month, 1);
                DateTime endMonth = combo.AddMonths(1);
                endMonth = endMonth.AddDays(-1);

                string bMonth = beginMonth.ToString("yyyy/MM/dd");
                string eMonth = endMonth.ToString("yyyy/MM/dd");

                thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT TimeIn, TimeOut FROM event WHERE CaseID = " + SelectCase3DDL.SelectedValue + " AND (TimeIn BETWEEN '" + bMonth + "' AND '" + eMonth + "' OR  TimeOut BETWEEN '" + bMonth + "' AND '" + eMonth + "');";
                //thisCommand.CommandText = "SELECT TimeIn, TimeOut FROM event WHERE CaseID = " + SelectCase3DDL.SelectedValue + " AND TimeIn BETWEEN '" + beginMonth + "' AND '" + endMonth + "' OR  TimeOut BETWEEN '" + beginMonth + "' AND '" + endMonth + "';";
                dataAdapter.SelectCommand = thisCommand;
                dataAdapter.Fill(DS, "MonthHours");

                thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT TimeIn, TimeOut FROM event WHERE CaseID = " + SelectCase3DDL.SelectedValue + ";";
                dataAdapter.SelectCommand = thisCommand;
                dataAdapter.Fill(DS, "AllHours");

                string LName = Convert.ToString(DS.Tables["PersonInfo"].Rows[0].ItemArray[2]);
                string FName = Convert.ToString(DS.Tables["PersonInfo"].Rows[0].ItemArray[0]);
                string MName = Convert.ToString(DS.Tables["PersonInfo"].Rows[0].ItemArray[1]);
                DateTime LastContact = Convert.ToDateTime(DS.Tables["PersonInfo"].Rows[0].ItemArray[4]);

                DateTime AgencyContactBy = Convert.ToDateTime(DS.Tables["CaseInfo"].Rows[0].ItemArray[2]);
                DateTime EstCompDate = Convert.ToDateTime(DS.Tables["CaseInfo"].Rows[0].ItemArray[3]);
                DateTime ReportMonth = dateTimePicker6.Value;

                double MHours = 0;
                DataTable temp = DS.Tables["MonthHours"];
                foreach (DataRow dr in temp.Rows)
                {
                    if (dr.ItemArray[0].ToString() == "0/0/0000 12:00:00 AM")
                    {
                        continue;
                    }
                    double hours;
                    DateTime tIn = Convert.ToDateTime(dr.ItemArray[0]);
                    DateTime tOut = Convert.ToDateTime(dr.ItemArray[1]);
                    TimeSpan ts = tOut - tIn;
                    hours = (ts.Days * 24) + ts.Hours + (ts.Minutes / 60);
                    //MHours += hours;
                    MHours += (Math.Round(ts.TotalHours, 2));
                }

                double CHours = 0;
                temp = DS.Tables["AllHours"];
                foreach (DataRow dr in temp.Rows)
                {
                    if (dr.ItemArray[0].ToString() == "0/0/0000 12:00:00 AM")
                    {
                        continue;
                    }
                    double hours;
                    DateTime tIn = Convert.ToDateTime(dr.ItemArray[0]);
                    DateTime tOut = Convert.ToDateTime(dr.ItemArray[1]);
                    TimeSpan ts = tOut - tIn;
                    hours = (ts.Days * 24) + ts.Hours + (ts.Minutes / 60);
                    //CHours += hours;
                    CHours += (Math.Round(ts.TotalHours, 2));
                }

                DateTime FirstFriday = new DateTime(combo.Year, combo.Month, 1);
                FirstFriday = FirstFriday.AddMonths(1);
                while (FirstFriday.Day != 5)
                {
                    FirstFriday = FirstFriday.AddDays(1);
                }

                app.Selection.Font.Size = 12;
                app.Selection.Font.Name = "Times New Roman";
                app.Selection.PageSetup.TopMargin = InchesToPoints((float)1);
                app.Selection.PageSetup.BottomMargin = InchesToPoints((float)1);
                app.Selection.PageSetup.LeftMargin = InchesToPoints((float)1.25);
                app.Selection.PageSetup.RightMargin = InchesToPoints((float)1.25);

                app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                app.Selection.TypeText(ReportMonth.ToString("MMMM") + ", " + ReportMonth.Year + " MONTHLY COMMUNITY SERVICE HOURS");
                app.Selection.TypeText("\n\n");
                app.Selection.TypeText("The following individual should be actively doing community service hours with your organization." +
                                        "  Please fill out the form and return it to our office at 721 Oxford Avenue, Eau Claire, WI" +
                                        " 54702 - Attn: Jessica, or FAX to 715-839-4817.  Thank you!");
                app.Selection.TypeText("\n");
                app.Selection.Font.Bold = 1;
                app.Selection.TypeText("Current Agency");
                app.Selection.Font.Bold = 0;
                app.Selection.TypeText("\t\t");
                app.Selection.TypeText("Feed My People Food Bank");
                app.Selection.TypeText("\n");
                app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                object autoFit1 = true;
                object autoFit2 = true;
                Table t = app.Selection.Tables.Add(app.Selection.Range, 2, 6, ref autoFit1, ref autoFit2);
                t.Borders.Enable = 0;
                t.Cell(1, 1).Range.Text = "Last Name";
                t.Cell(1, 1).Range.Font.Bold = 1;
                t.Cell(1, 2).Range.Text = "First Name";
                t.Cell(1, 2).Range.Font.Bold = 1;
                t.Cell(1, 3).Range.Text = "M";
                t.Cell(1, 3).Range.Font.Bold = 1;
                t.Cell(1, 3).Column.SetWidth(20, WdRulerStyle.wdAdjustSameWidth);
                t.Cell(1, 4).Range.Text = ReportMonth.ToString("MMMM") + " Hours";
                t.Cell(1, 4).Range.Font.Bold = 1;
                t.Cell(1, 5).Range.Text = "Completed Hours";
                t.Cell(1, 5).Range.Font.Bold = 1;
                t.Cell(1, 6).Range.Text = "Last Contact Date";
                t.Cell(1, 6).Range.Font.Bold = 1;

                t.Cell(2, 1).Range.Text = LName;
                t.Cell(2, 2).Range.Text = FName;
                t.Cell(2, 3).Range.Text = MName;
                t.Cell(2, 4).Range.Text = MHours.ToString();
                t.Cell(2, 5).Range.Text = CHours.ToString();
                t.Cell(2, 6).Range.Text = LastContact.ToShortDateString();
                t.Select();
                object move = WdUnits.wdLine;
                object move2 = 1;
                object move3 = WdMovementType.wdMove;
                app.Selection.MoveDown(ref move, ref move2, ref move3);

                app.Selection.TypeText("\n\n");
                app.Selection.Font.Bold = 1;
                app.Selection.TypeText("Additional Notes/Comments:");
                app.Selection.Font.Bold = 0;
                app.Selection.TypeText("\n");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                app.Selection.TypeText("_______________________________________________________________________\n");
                app.Selection.TypeText("_______________________________________________________________________\n");
                app.Selection.TypeText("_______________________________________________________________________\n");
                app.Selection.TypeText("_______________________________________________________________________\n");
                app.Selection.TypeText("_______________________________________________________________________\n");
                app.Selection.TypeText("_______________________________________________________________________\n");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                app.Selection.TypeText("\n");
                app.Selection.Font.Bold = 1;
                app.Selection.TypeText("Notes: ");
                app.Selection.TypeText("\n");
                app.Selection.TypeText("Est Completion Date:");
                app.Selection.TypeText("\t");
                app.Selection.Font.Bold = 0;
                app.Selection.TypeText(EstCompDate.ToShortDateString());
                app.Selection.Font.Bold = 1;
                app.Selection.TypeText("\t\t");
                app.Selection.TypeText("Agency contact by:");
                app.Selection.Font.Bold = 0;
                app.Selection.TypeText("\t");
                app.Selection.TypeText(AgencyContactBy.ToShortDateString());
                app.Selection.Font.Bold = 1;
                app.Selection.TypeText("\n");
                app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                app.Selection.TypeText("Please return no later than " + FirstFriday.ToString("dddd") + ", " + FirstFriday.ToString("MMMM") + " " + FirstFriday.Day.ToString() + ", " + FirstFriday.Year + ".  Thank you.");
                app.Selection.Font.Bold = 0;
                app.Selection.TypeText("\n");

                Table t2 = app.Selection.Tables.Add(app.Selection.Range, 2, 2, ref autoFit1, ref autoFit2);
                t2.Borders.Enable = 0;
                app.Selection.Tables[1].Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                t2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                t2.Cell(1, 1).Column.SetWidth(200, WdRulerStyle.wdAdjustSameWidth);
                t2.Cell(1, 1).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                t2.Cell(1, 1).Range.Text = "_____________________";
                t2.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                t2.Cell(1, 2).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                t2.Cell(1, 2).Range.Text = "__________";
                t2.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                t2.Cell(1, 2).Column.SetWidth(80, WdRulerStyle.wdAdjustSameWidth);
                t2.Cell(2, 1).Range.Text = "Supervisor Signature";
                t2.Cell(2, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                t2.Cell(2, 2).Range.Text = "Date";
                t2.Cell(2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                t2.Select();
                app.Selection.MoveDown(ref move, ref move2, ref move3);

                app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                app.Visible = true;
            }
            catch (MySqlException ee)
            {
                MessageBox.Show("An error occured connecting to the database!");
            }
            catch (Exception eee)
            {
                MessageBox.Show(eee.Message);
            }
            finally
            {
                thisConnection.Close();
            }

        }

        private float InchesToPoints(float inches)
        {
            return (inches * 72);
        }
        private void CreateTimeSheetReportButton_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.ApplicationClass app = new Microsoft.Office.Interop.Word.ApplicationClass();
                Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();

                object missing = System.Reflection.Missing.Value;
                object b = false;
                object c = 0;
                object d = false;
                doc = app.Documents.Add(ref missing, ref b, ref c, ref d);
                doc.Activate();

                string connectionString = ConfigurationManager.ConnectionStrings["DB"].ConnectionString;
                thisConnection = new MySqlConnection(connectionString);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = thisCommand.CommandText = "SELECT CaseNum, HoursAssigned, AgencyContactDate, EstimatedCompletionDate, Status, WeeklyReq FROM cases WHERE PersonID = '" + SelectVolunteer3DDL.SelectedValue + "' AND CaseID = '" + SelectCase3DDL.SelectedValue + "';";
                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);

                DataSet DS = new DataSet();
                dataAdapter.Fill(DS, "CaseInfo");

                thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT FName, MInitial, LName, Address, LastContactDay, Status FROM person WHERE PersonID = '" + SelectVolunteer3DDL.SelectedValue + "';";
                dataAdapter.SelectCommand = thisCommand;
                dataAdapter.Fill(DS, "PersonInfo");

                thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT TimeIn, TimeOut FROM event WHERE CaseID = " + SelectCase3DDL.SelectedValue + ";";
                dataAdapter.SelectCommand = thisCommand;
                dataAdapter.Fill(DS, "AllHours");

                string vName = Convert.ToString(DS.Tables["PersonInfo"].Rows[0].ItemArray[0]) + " " + Convert.ToString(DS.Tables["PersonInfo"].Rows[0].ItemArray[1]) + " " + Convert.ToString(DS.Tables["PersonInfo"].Rows[0].ItemArray[2]);
                double HoursAssigned = Convert.ToDouble(DS.Tables["CaseInfo"].Rows[0].ItemArray[1]);
                string caseNum = Convert.ToString(DS.Tables["CaseInfo"].Rows[0].ItemArray[0]);
                double weeklyReq = Convert.ToInt32(DS.Tables["CaseInfo"].Rows[0].ItemArray[5]);
                double totalHours = 0.0;
                DateTime completionDate = DateTime.MinValue;
                DateTime reportDate = DateTime.Now;
                string agency = "Feed My People Food Bank";
                int numRows = 0;

                DataTable temp = DS.Tables["AllHours"];
                foreach (DataRow dr in temp.Rows)
                {
                    if (dr.ItemArray[0].ToString() == "0/0/0000 12:00:00 AM")
                    {
                        continue;
                    }
                    double hours;
                    DateTime tIn = Convert.ToDateTime(dr.ItemArray[0]);
                    DateTime tOut = Convert.ToDateTime(dr.ItemArray[1]);
                    TimeSpan ts = tOut - tIn;
                    if (ts.Minutes == 0 && ts.Hours == 0 && ts.Days == 0)
                    {
                        continue;
                    }
                    hours = (ts.Days * 24) + ts.Hours + (ts.Minutes / 60);
                    //totalHours += hours;
                    totalHours += (Math.Round(ts.TotalHours, 2));
                    if (tOut > completionDate)
                    {
                        completionDate = Convert.ToDateTime(dr.ItemArray[1]);
                    }
                    numRows++;
                }

                object autoFit1 = true;
                object autoFit2 = true;

                app.Selection.Font.Size = 12;
                app.Selection.Font.Name = "Times New Roman";
                app.Selection.PageSetup.TopMargin = InchesToPoints((float)1);
                app.Selection.PageSetup.BottomMargin = InchesToPoints((float)1);
                app.Selection.PageSetup.LeftMargin = InchesToPoints((float)1.25);
                app.Selection.PageSetup.RightMargin = InchesToPoints((float)1.25);

                Table t3 = app.Selection.Tables.Add(app.Selection.Range, 2, 2, ref autoFit1, ref autoFit2);
                t3.Borders.Enable = 0;
                t3.Cell(1, 1).Column.SetWidth(330, WdRulerStyle.wdAdjustSameWidth);
                t3.Cell(1, 1).Range.Text = "Eau Claire County Community Service Program";
                t3.Cell(1, 2).Column.SetWidth(160, WdRulerStyle.wdAdjustSameWidth);
                t3.Cell(1, 2).Range.Text = "Return to: Community Service";
                t3.Cell(2, 1).Range.Text = "Participant Time Sheet";
                t3.Cell(2, 2).Range.Text = "Clerk of Courts, 721 Oxford Avenue\n" +
                                            "Eau Claire, WI 54703\n" +
                                            "FAX: 715-839-4817\n";

                t3.Select();
                object move = WdUnits.wdLine;
                object move2 = 1;
                object move3 = WdMovementType.wdMove;
                app.Selection.MoveDown(ref move, ref move2, ref move3);

                app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                //string under1 = "                                                                           ";
                //string under2 = "             ";
                string under1 = "                                             ";
                string under2 = "      ";
                string under3 = "";
                app.Selection.TypeText("Name:  ");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                app.Selection.TypeText(vName + under1.Substring(0, under1.Length - vName.Length));
                app.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                app.Selection.TypeText("     Number of Hours Assigned:  ");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                app.Selection.TypeText(HoursAssigned.ToString() + under2.Substring(0, under2.Length - HoursAssigned.ToString().Length) + "_");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                app.Selection.TypeText("\n");

                under1 = "                 ";
                under2 = "        ";
                under3 = "              ";

                app.Selection.TypeText("Case #:  ");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                app.Selection.TypeText(caseNum.ToString() + under1.Substring(0, under1.Length - caseNum.ToString().Length));
                app.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                app.Selection.TypeText("   Weekly Requirement:  ");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                app.Selection.TypeText(weeklyReq.ToString() + under2.Substring(0, under2.Length - weeklyReq.ToString().Length));
                app.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                app.Selection.TypeText("   Completion Date:  ");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                app.Selection.TypeText(completionDate.ToShortDateString() + under3.Substring(0, under3.Length - completionDate.ToShortDateString().Length) + "_");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                app.Selection.TypeText("\n");

                numRows++; //add one more to rows for the header
                int numOfCols = 5;  //should always be the same but can change
                Table t = app.Selection.Tables.Add(app.Selection.Range, numRows, numOfCols, ref autoFit1, ref autoFit2);
                t.Borders.Enable = 1;

                t.Cell(1, 1).Column.SetWidth(60, WdRulerStyle.wdAdjustSameWidth);
                t.Cell(1, 1).Range.Text = "Date";
                t.Cell(1, 2).Column.SetWidth(60, WdRulerStyle.wdAdjustSameWidth);
                t.Cell(1, 2).Range.Text = "Time In";
                t.Cell(1, 3).Column.SetWidth(60, WdRulerStyle.wdAdjustSameWidth);
                t.Cell(1, 3).Range.Text = "Time Out";
                t.Cell(1, 4).Column.SetWidth(80, WdRulerStyle.wdAdjustSameWidth);
                t.Cell(1, 4).Range.Text = "Total Hours";
                t.Cell(1, 5).Column.SetWidth(220, WdRulerStyle.wdAdjustSameWidth);
                t.Cell(1, 5).Range.Text = "CALL 839-1869 by first Friday of EVERY MONTH with report of hours completed";
                int rowCount = 2;
                foreach (DataRow dr in temp.Rows)
                {
                    if (dr.ItemArray[0].ToString() == "0/0/0000 12:00:00 AM")
                    {
                        continue;
                    }
                    //double hours;
                    DateTime tIn = Convert.ToDateTime(dr.ItemArray[0]);
                    DateTime tOut = Convert.ToDateTime(dr.ItemArray[1]);
                    TimeSpan ts = tOut - tIn;
                    if (ts.Minutes == 0 && ts.Hours == 0 && ts.Days == 0)
                    {
                        continue;
                    }
                    //hours = (ts.Days * 24) + ts.Hours + (ts.Minutes / 60);

                    t.Cell(rowCount, 1).Range.Text = tIn.ToShortDateString();
                 //   t.Cell(rowCount, 2).Range.Text = tIn.ToString("MM/dd/yyyy") + " " + tIn.ToString("hh:mm") + " " + tIn.ToString("tt");
                 //   t.Cell(rowCount, 3).Range.Text = tOut.ToString("MM/dd/yyyy") + " " + tOut.ToString("hh:mm") + " " + tOut.ToString("tt");

                    t.Cell(rowCount, 2).Range.Text = tIn.ToString("hh:mm") + " " + tIn.ToString("tt");
                    t.Cell(rowCount, 3).Range.Text = tOut.ToString("hh:mm") + " " + tOut.ToString("tt");


                    //t.Cell(rowCount, 4).Range.Text = hours.ToString();
                    t.Cell(rowCount, 4).Range.Text = (Math.Round(ts.TotalHours, 2)).ToString();
                    rowCount++;
                }

                t.Select();
                app.Selection.MoveDown(ref move, ref move2, ref move3);

                app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                app.Selection.TypeText("Report of Completed Community Service");
                app.Selection.TypeText("\n");

                under1 = "                                                   ";
                under2 = "        ";
                app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                app.Selection.TypeText("I hereby certify that ");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                app.Selection.TypeText(vName + under1.Substring(0, under1.Length - vName.ToString().Length));
                app.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                app.Selection.TypeText(" has satisfactorily completed ");
                app.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                app.Selection.TypeText(totalHours + under2.Substring(0, totalHours.ToString().Length));
                app.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
                app.Selection.TypeText(" hours of community service\n");

                Table t2 = app.Selection.Tables.Add(app.Selection.Range, 2, 3, ref autoFit1, ref autoFit2);
                t2.Borders.Enable = 0;

                under1 = "               ";
                under2 = "                                       ";
                under3 = "                      ";
                t2.Cell(1, 1).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                t2.Cell(1, 1).Range.Text = reportDate.ToShortDateString() + under1.Substring(0, reportDate.ToShortDateString().Length);
                t2.Cell(1, 2).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                t2.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                t2.Cell(1, 2).Range.Text = "_______________________________________";
                t2.Cell(1, 3).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                t2.Cell(1, 3).Range.Text = agency;
                t2.Cell(2, 1).Range.Text = "Date";
                t2.Cell(2, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                t2.Cell(2, 1).Column.SetWidth(35, WdRulerStyle.wdAdjustSameWidth);
                t2.Cell(2, 2).Range.Text = "Site Supervisor";
                t2.Cell(2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                t2.Cell(2, 2).Column.SetWidth(290, WdRulerStyle.wdAdjustSameWidth);
                t2.Cell(2, 3).Range.Text = "Agency";
                t2.Cell(2, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                t2.Cell(2, 3).Column.SetWidth(160, WdRulerStyle.wdAdjustSameWidth);

                t2.Select();
                app.Selection.MoveDown(ref move, ref move2, ref move3);

                app.Visible = true;
            }
            catch (MySqlException ee)
            {
                MessageBox.Show("An error occured connecting to the database!");
            }
            catch (Exception eee)
            {
                MessageBox.Show(eee.Message);
            }
            finally
            {
                thisConnection.Close();
            }      
        }

        //Cancel changes in progress if user changes tab
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (AddNewVolunteerButton.Text.Equals("Cancel Add New Volunteer"))
            {
                isUpdate = true;
                SelectVolunteer2DDL.Enabled = true;
                AddNewVolunteerButton.Text = "Add New Volunteer";
                setAddNewVolunteerButtons(true);
                FillVolunteerInfoFields();
            }

            if (EditCaseButton.Text.Equals("Cancel Editing Case"))
            {
                enableCase(false);
                setEditCaseButtons(true);
                SaveNewCaseButton.Text = "Save";
                EditCaseButton.Text = "Edit Selected Case";
            }

            if (btnAddCase.Text.Equals("Cancel Add Case"))
            {
                enableCase(false);
                setAddNewCaseButtons(true);
                btnAddCase.Text = "Add New Case";
                FillCaseInfoFields();
            }

            //Reset the enter time tab
            ResetEnterTime();
        }

        private void ResetEnterTime()
        {
            resettingTime = true;
            DateTime time = DateTime.Now;
            //constructor as follows:
            //DateTime(int year, int month, int day, int hour, int minute, int second, int millisecond);
            //create these dates by hand so that the milliseconds don't mess up calculations (set milliseconds on both to zero)
            DateTime now = new DateTime(time.Year, time.Month, time.Day, time.Hour,
                time.Minute, time.Second, 0);

            if (now.Hour >= 12)
            {   //PM
                string hour;
                if (now.Hour == 12)
                {
                    hour = "12";
                }
                else
                {
                    hour = Convert.ToString(now.Hour - 12);
                }
                inHour.Text = hour;
                outHour.Text = hour;
                //set to PM
                inAmPm.SelectedIndex = 1;
                outAmPm.SelectedIndex = 1;
            }
            else
            {   //AM
                string hour;
                if (now.Hour == 0)
                {
                    hour = "12";
                }
                else
                {
                    hour = Convert.ToString(now.Hour);
                }
                inHour.Text = hour;
                outHour.Text = hour;
                //set to AM
                inAmPm.SelectedIndex = 0;
                outAmPm.SelectedIndex = 0;
            }

            if (now.Minute < 10 && now.Minute >= 0)
            {
                inMin.Text = "0" + now.Minute;
                outMin.Text = "0" + now.Minute;
            }
            else
            {
                inMin.Text = Convert.ToString(now.Minute);
                outMin.Text = Convert.ToString(now.Minute);
            }

            inMonth.Text = Convert.ToString(now.Month);
            inDay.Text = Convert.ToString(now.Day);
            inYear.Text = Convert.ToString(now.Year);

            outMonth.Text = Convert.ToString(now.Month);
            outDay.Text = Convert.ToString(now.Day);
            outYear.Text = Convert.ToString(now.Year);

            TotalHoursTextBox.Text = "00:00";
            resettingTime = false;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void SelectCase1DDL_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtTotalHours.Text = String.Format("{0}:{1}","00","00");
            txtNeeded.Text = String.Format("{0}:{1}", "00", "00");
            grdViewHours.DataSource = null;
            if (SelectCase1DDL.Items.Count > 0)
            {
                if (SelectCase1DDL.SelectedValue != null && !SelectCase1DDL.SelectedValue.Equals(null))
                {
                    SelectCase1DDL.SelectedIndex = SelectCase1DDL.FindString(SelectCase1DDL.Text);
                }
                fillTimeGrid();
            }
        }

        private void fillTimeGrid()
        {
            try
            {
                DataSet DS = new DataSet();

                string connectionString = ConfigurationManager.ConnectionStrings["DB"].ConnectionString;
                thisConnection = new MySqlConnection(connectionString);
                thisConnection.Open();
                MySqlCommand thisCommand = thisConnection.CreateCommand();
                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(thisCommand.CommandText, thisConnection);
                MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(dataAdapter);
                thisCommand = thisConnection.CreateCommand();
                thisCommand.CommandText = "SELECT TimeIn, TimeOut FROM event WHERE CaseID = '" + SelectCase1DDL.SelectedValue + "';";
                dataAdapter.SelectCommand = thisCommand;
                dataAdapter.Fill(DS, "AllHours");

                thisCommand.CommandText = "SELECT HoursAssigned, WeeklyReq FROM cases WHERE PersonID = '" + SelectVolunteer1DDL.SelectedValue + "' AND CaseID = '" + SelectCase1DDL.SelectedValue + "';";
                dataAdapter.SelectCommand = thisCommand;
                dataAdapter.Fill(DS, "PersonHours");

                string totalHours = "0";
                string totalMins = "0";
                DS.Tables["AllHours"].Columns.Add("Time Difference", System.Type.GetType("System.String"));
                foreach(DataRow r in DS.Tables["AllHours"].Rows){
                    string hours = "0";
                    string mins = "0";
                    //constructor as follows:
                    //DateTime(int year, int month, int day, int hour, int minute, int second, int millisecond);
                    //create these dates by hand so that the milliseconds don't mess up calculations (set milliseconds on both to zero)
                    DateTime initialTimeIn = Convert.ToDateTime(r["TimeIn"]);
                    DateTime initialTimeOut = Convert.ToDateTime(r["TimeOut"]);

                    DateTime timeIn = new DateTime(initialTimeIn.Year, initialTimeIn.Month, initialTimeIn.Day, initialTimeIn.Hour,
                        initialTimeIn.Minute, initialTimeIn.Second, 0);
                    DateTime timeOut = new DateTime(initialTimeOut.Year, initialTimeOut.Month, initialTimeOut.Day, initialTimeOut.Hour,
                        initialTimeOut.Minute, initialTimeOut.Second, 0);
                    
                    TimeSpan ts = timeOut - timeIn;

                    //update totals:
                    totalHours = Convert.ToString(Convert.ToInt32(totalHours) + ts.Hours);
                    totalMins = Convert.ToString(Convert.ToInt32(totalMins) + ts.Minutes);

                    hours = Convert.ToInt32(ts.Hours).ToString();
                    mins = Convert.ToInt32(ts.Minutes).ToString();
                    if ((Convert.ToInt32(mins) < 10) && (Convert.ToInt32(mins) >= 0))
                    {
                        mins = "0" + mins;
                    }
                    if ((Convert.ToInt32(hours) < 10) && (Convert.ToInt32(hours) >= 0))
                    {
                        hours = "0" + hours;
                    }

                    r["Time Difference"] = String.Format("{0}:{1}", hours, mins);
                }
                if (Convert.ToInt32(totalMins) > 59)
                {
                    int extraHours = Convert.ToInt32(totalMins) / 60;
                    totalHours = Convert.ToString(Convert.ToInt32(totalHours) + extraHours);
                    totalMins = Convert.ToString(Convert.ToInt32(totalMins) % 60);
                }
                if ((Convert.ToInt32(totalMins) < 10) && (Convert.ToInt32(totalMins) >= 0))
                {
                    totalMins = "0" + totalMins;
                }
                if ((Convert.ToInt32(totalHours) < 10) && (Convert.ToInt32(totalHours) >= 0))
                {
                    totalHours = "0" + totalHours;
                }
                txtTotalHours.Text = String.Format("{0}:{1}", totalHours, totalMins);

                if (DS.Tables["AllHours"].Rows.Count > 0)
                {
                    grdViewHours.DataSource = DS.Tables["AllHours"];
                }

                DataRow dr = DS.Tables["PersonHours"].Rows[0];
                string neededHours = Convert.ToString(dr["HoursAssigned"]);
                string hrs;
                string minutes;
                if (neededHours.Contains('.'))
                {
                    char[] sep = { '.' };
                    string[] splitHours = neededHours.Split(sep, 2);
                    hrs = splitHours[0];
                    minutes = splitHours[1];
                }
                else
                {
                    hrs = neededHours;
                    minutes = "0";
                }

                if ((Convert.ToInt32(minutes) < 10) && (Convert.ToInt32(minutes) >= 0))
                {
                    minutes = "0" + minutes;
                }
                if ((Convert.ToInt32(hrs) < 10) && (Convert.ToInt32(hrs) >= 0))
                {
                    hrs = "0" + hrs;
                }

                txtNeeded.Text = String.Format("{0}:{1}", hrs, minutes);
            }
            catch (MySqlException ee)
            {
                MessageBox.Show("An error occured connecting to the database!");
            }
            catch (Exception eee)
            {
                //MessageBox.Show(eee.Message);
            }
            finally
            {
                thisConnection.Close();
            }  
        }

        private DateTime getTimeIn()
        {
            int hour;
            if (inAmPm.SelectedIndex == 0)
            {   //AM
                if (inHour.Text.Equals("12"))
                {
                    hour = 0;
                }
                else
                {
                    hour = Convert.ToInt32(inHour.Text);
                }
            }
            else
            {   //PM
                if (inHour.Text.Equals("12"))
                {
                    hour = 12;
                }
                else
                {
                    hour = Convert.ToInt32(inHour.Text) + 12;
                }
            }
            //constructor takes year, month, day, hour, minute, second
            DateTime inDate = new DateTime(Convert.ToInt32(inYear.Text), Convert.ToInt32(inMonth.Text), Convert.ToInt32(inDay.Text),
                hour, Convert.ToInt32(inMin.Text), 0);
            return inDate;
        }

        private DateTime getTimeOut()
        {
            int hour;
            if (outAmPm.SelectedIndex == 0)
            {   //AM
                if (outHour.Text.Equals("12"))
                {
                    hour = 0;
                }
                else
                {
                    hour = Convert.ToInt32(outHour.Text);
                }
            }
            else
            {   //PM
                if (outHour.Text.Equals("12"))
                {
                    hour = 12;
                }
                else
                {
                    hour = Convert.ToInt32(outHour.Text) + 12;
                }
            }
            //constructor takes year, month, day, hour, minute, second
            DateTime outDate = new DateTime(Convert.ToInt32(outYear.Text), Convert.ToInt32(outMonth.Text), Convert.ToInt32(outDay.Text),
                hour, Convert.ToInt32(outMin.Text), 0);
            return outDate;
        }

        private void updateEvent(object sender, EventArgs e)
        {
            updateTimeIn(grdViewHours.CurrentRow.Cells[0].Value.ToString());
            updateTimeOut(grdViewHours.CurrentRow.Cells[1].Value.ToString());
        }

        private void updateTimeIn(string cell) //cell is in format "DD/MM/YYYY HH:MM:SS PM"
        {
            string date = cell.Split(' ')[0];
            string time = cell.Split(' ')[1];
            string amPm = cell.Split(' ')[2];
            inMonth.Text = date.Split('/')[0];
            inDay.Text = date.Split('/')[1];
            inYear.Text = date.Split('/')[2];
            inHour.Text = time.Split(':')[0];
            inMin.Text = time.Split(':')[1];
            inAmPm.Text = amPm;
        }

        private void updateTimeOut(string cell) //cell is in format "DD/MM/YYYY HH:MM:SS PM"
        {
            string date = cell.Split(' ')[0];
            string time = cell.Split(' ')[1];
            string amPm = cell.Split(' ')[2];
            outMonth.Text = date.Split('/')[0];
            outDay.Text = date.Split('/')[1];
            outYear.Text = date.Split('/')[2];
            outHour.Text = time.Split(':')[0];
            outMin.Text = time.Split(':')[1];
            outAmPm.Text = amPm;
        }

        private void dataGridCellClick(object sender, DataGridViewCellEventArgs e)
        {
            updateButton.Enabled = true;
            selectedColumnIndex = e.ColumnIndex;
            selectedRowIndex = e.RowIndex;
        }
    }
}
