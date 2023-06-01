using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Data.OleDb;
using Excel;
using System.Diagnostics;

namespace ProjectCIS
{
    public partial class frmMain : Form
    {
        System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("th-TH");
        string homePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
        OleDbCommand cmd = new OleDbCommand();
        OleDbConnection con = new OleDbConnection();
        OleDbDataReader dr;
        OpenFileDialog ope;
        string idInfo;
        string filePath;
        int checkLoad = 0; //ไว้ตรวจสอบตอนอัพโหลดไฟล์ 0=ปกติ 1=ตรวจสอบก่อนโหลดไฟล์เสร็จ 2=ไม่ให้ปิดฟอร์ม
        string checkBA;
        string checkBAReg;
        string lblUnitBA;
        string lblUnitBAReg = "";
        string checkBoxReview;
        string[] filesOpen;

        [Flags] enum checkFileStatus
        {
            none = 0x0, //ค่าว่าง
            allCustomer = 0x1, //ผู้ใช้น้ำทั้งหมด
            installCost = 0x2, //การรับเงินค่าติดตั้ง
            waterRevenue = 0x4, //การรับเงินรายได้ค่าน้ำ
            otherRevenueCustomer = 0x8, //รายได้อื่นๆเกี่ยวกับผู้ใช้น้ำ
            otherRevenueNonCustomer = 0x10, //รายได้อื่นๆไม่เกี่ยวกับผู้ใช้น้ำ
            depositMeter = 0x20, //ฝากมาตรครบกำหนด
            pipeMeter = 0x40, //ทะเบียนคุมท่อธาร
            debtMonth = 0x80, //ตั้งหนี้ประจำเดือน
            debtCustomer = 0x100,
            garuntee = 0x200, //หลักประกันสัญญา
            debtCurrent = 0x400, //หนี้ค้างชำระ
            meterHistory = 0x800, //ประวัติการเปลี่ยนสถานะมาตร
            meterAbnormal = 0x1000, //มาตรตาย
            cancelReceipt = 0x2000, //ยกเลิกใบสร็จ
        };
        [Flags] enum checkFileStatusReg
        {
            none = 0x0, //ค่าว่าง
            installCost = 0x1, //การรับเงินค่าติดตั้ง
            waterRevenue = 0x2, //การรับเงินรายได้ค่าน้ำ
            otherRevenueNonCustomer = 0x4, //รายได้อื่นๆไม่เกี่ยวกับผู้ใช้น้ำ
            garuntee = 0x8, //หลักประกันสัญญา
            cancelReceipt = 0x10, //ยกเลิกใบสร็จ
            debtLitigate = 0x20, //หนี้อยู่ระหว่างดำเนินคดี
        }
        //กำหนดค่า enum เริ่มต้นให้ flag เป็น none
        checkFileStatus fileStatusFlag = checkFileStatus.none;
        //กำหนดค่า flag ให้แต่ละปุ่มกระดาษทำการ
        checkFileStatus[] btnEnable = {checkFileStatus.allCustomer , checkFileStatus.installCost , checkFileStatus.waterRevenue ,
                                        checkFileStatus.otherRevenueCustomer , checkFileStatus.otherRevenueNonCustomer , checkFileStatus.depositMeter ,
                                        checkFileStatus.pipeMeter , checkFileStatus.debtMonth , checkFileStatus.debtCustomer , checkFileStatus.garuntee ,
                                        checkFileStatus.debtCurrent , checkFileStatus.meterHistory , checkFileStatus.meterAbnormal , checkFileStatus.cancelReceipt };
        //ตัวอย่างใช้สองไฟล์  checkFileStatus.allCustomer|checkFileStatus.installCost

        //กำหนดค่า enum เริ่มต้นให้ flag เป็น none
        checkFileStatusReg fileStatusFlagReg = checkFileStatusReg.none;
        //กำหนดค่า flag ให้แต่ละปุ่มกระดาษทำการ
        checkFileStatusReg[] btnEnableReg = {checkFileStatusReg.installCost , checkFileStatusReg.waterRevenue , checkFileStatusReg.otherRevenueNonCustomer ,
                                              checkFileStatusReg.garuntee , checkFileStatusReg.cancelReceipt , checkFileStatusReg.debtLitigate };

        public frmMain()
        {
            InitializeComponent();
            System.Threading.Thread.CurrentThread.CurrentCulture = cultureInfo;
            System.Threading.Thread.CurrentThread.CurrentUICulture = cultureInfo;

        }
        
        //เครดิต
        private void toolStripButtonCredit_Click(object sender, EventArgs e)
        {
            MessageBox.Show("นายชัชพล นุโยค\nนายเอกพจน์ ศักดิ์เรืองฤทธิ์\nนายศิริชัย แก่นไชย\nVersion 1.20 วันที่ 28 ตุลาคม 2565","ผู้พัฒนา",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        //ตั้งค่าเปิด/ปิดปุ่มกระดาษทำการสาขา
        private void setButton()
        {
          
            if (fileStatusFlag.HasFlag(btnEnable[0]))
            {
                btnRandomMeter.Enabled = true;
                btnDiscount.Enabled = true;
            }
            else
            {
                btnRandomMeter.Enabled = false;          
                btnDiscount.Enabled = false;
                
            }

            if (fileStatusFlag.HasFlag(btnEnable[1]))
            {
                btnInstallCost.Enabled = true;
                btnRegisterCustomer.Enabled = true;
            }
            else
            {
                btnInstallCost.Enabled = false;
                btnRegisterCustomer.Enabled = false;

            }

            if (fileStatusFlag.HasFlag(btnEnable[2]))
            {
                btnReconcileMoney.Enabled = true;
                btnConfirmBalance.Enabled = true; 
            }
            else
            {
                btnReconcileMoney.Enabled = false;
                btnConfirmBalance.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[3]))
            {
                btnOtherRevenueCustomer.Enabled = true;
            }
            else
            {
                btnOtherRevenueCustomer.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[4]))
            {
                btnOtherRevenueNonCustomer.Enabled = true;
            }
            else
            {
                btnOtherRevenueNonCustomer.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[5]))
            {
                btnDepositMeter.Enabled = true;
            }
            else
            {
                btnDepositMeter.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[6]))
            {
                btnPipe.Enabled = true;
            }
            else
            {
                btnPipe.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[7]))
            {
                btnCalWater.Enabled = true;
                btnAdjustDebt.Enabled = true;
            }
            else
            {
                btnCalWater.Enabled = false;
                btnAdjustDebt.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[8]))
            {

            }
            else
            {

            }

            if (fileStatusFlag.HasFlag(btnEnable[9]))
            {
                btnGaruntee.Enabled = true;
            }
            else
            {
                btnGaruntee.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[10]))
            {
                btnDebtCurrent.Enabled = true;
            }
            else
            {
                btnDebtCurrent.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[11]))
            {
                
            }
            else
            {

            }

            if (fileStatusFlag.HasFlag(btnEnable[12]))
            {
                btnMeterAbnormal.Enabled = true;
            }
            else
            {
                btnMeterAbnormal.Enabled = false;
            }

            if (fileStatusFlag.HasFlag(btnEnable[13]))
            {
                btnCancelReceipt.Enabled = true;
            }
            else
            {
                btnCancelReceipt.Enabled = false;
            }


        }

        //ตั้งค่าเปิด/ปิดปุ่มกระดาษทำการเขต
        private void setButtonReg()
        {
            if (fileStatusFlagReg.HasFlag(btnEnableReg[0]))
            {
                btnInstallCostReg.Enabled = true;
            }
            else
            {
                btnInstallCostReg.Enabled = false;
            }

            if (fileStatusFlagReg.HasFlag(btnEnableReg[1]))
            {
                btnWaterRevenueReg.Enabled = true;
            }
            else
            {
                btnWaterRevenueReg.Enabled = false;
            }

            if (fileStatusFlagReg.HasFlag(btnEnableReg[2]))
            {
                btnOtherRevenueNonCustomerReg.Enabled = true;
            }
            else
            {
                btnOtherRevenueNonCustomerReg.Enabled = false;
            }

            if (fileStatusFlagReg.HasFlag(btnEnableReg[3]))
            {
                btnGarunteeReg.Enabled = true;
            }
            else
            {
                btnGarunteeReg.Enabled = false;
            }

            if (fileStatusFlagReg.HasFlag(btnEnableReg[4]))
            {
                btnCancelReceiptReg.Enabled = true;
            }
            else
            {
                btnCancelReceiptReg.Enabled = false;
            }

            if (fileStatusFlagReg.HasFlag(btnEnableReg[5]))
            {
                btnLitigateAll.Enabled = true;
            }
            else
            {
                btnLitigateAll.Enabled = false;
            }

        }

            //โหลดข้อมูลจากฐานข้อมูล Info
            private void loadInfo ()
        {
            try
            {
                con.Open();
                cmd.Connection = con;
                string query = "SELECT * FROM info";
                cmd.CommandText = query;
                dr = cmd.ExecuteReader();
                dr.Read();
                idInfo = dr["ID"].ToString();
                txtUnitName.Text = dr["unitName"].ToString();
                txtAuditName.Text = dr["auditName"].ToString();
                txtReviewName.Text = dr["reviewName"].ToString();
                dtpPeriodBegin.Value = Convert.ToDateTime(dr["periodBegin"]);
                dtpPeriodEnd.Value = Convert.ToDateTime(dr["periodEnd"]);
                dtpAuditDate.Value = Convert.ToDateTime(dr["auditDate"]);
                string regCheck = dr["checkReg"].ToString();

                /*
                dtpPeriodBegin.Format = DateTimePickerFormat.Custom;
                string[] formats = dtpPeriodBegin.Value.GetDateTimeFormats();
                dtpPeriodBegin.CustomFormat = formats[7];
                */
                panelMain.Show();
                panelWorkingPaper.Show();
                txtUnitName.Show();
                if (dr["reviewDate"].ToString() == "")
                {
                    dtpReviewDate.Enabled = false;
                    checkBoxReviewDate.Checked = true;
                }
                else
                {
                    dtpReviewDate.Value = Convert.ToDateTime(dr["reviewDate"]);
                }
                if (regCheck == "1")
                {
                    checkBoxUnitName.Checked = true;
                }
                else
                {
                    panelMainReg.Hide();
                    panelWorkingPaperReg.Hide();
                    comboBoxUnit.Hide();
                }
                con.Close();
                dr.Close();
            }
            catch (OleDbException)
            {
                MessageBox.Show("ไม่สามารถติดต่อฐานข้อมูลผู้ใช้งานได้", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                con.Close();
                dr.Close();
            }
        }
        //โหลดข้อมูลในฐานข้อมูลผู้ใช้น้ำทั้งหมด
        private void loadCustomerDB()
        {
           try
            {
                con.Open();
                cmd.Connection = con;
                string selectlbl = "SELECT top 1 ba.เลขที่สาขา, ba.ORG_OWNER_ID FROM ba INNER JOIN customer ON ba.ORG_OWNER_ID = customer.รหัสหน่วยงาน;";
                cmd.CommandText = selectlbl;
                dr = cmd.ExecuteReader();
                dr.Read();
                toolStripLabelUnit.Text = "BA " + dr[0].ToString();
                lblUnitBA = toolStripLabelUnit.Text;

                if (checkBoxUnitName.Checked)
                {
                    toolStripLabelUnit.Text = lblUnitBAReg;
                }
          
                lblStatusAllCustomer.Text = "R";
                lblStatusAllCustomer.ForeColor = Color.LightSeaGreen;
                checkBA = dr[1].ToString();
                fileStatusFlag |= checkFileStatus.allCustomer; //เซ็ตบิต
                con.Close();
                dr.Close();
            }
            catch
            {
                toolStripLabelUnit.Text = "ไม่พบฐานข้อมูลผู้ใช้น้ำ";
                lblStatusAllCustomer.Text = "ฃ";
                lblStatusAllCustomer.ForeColor = SystemColors.ControlDarkDark;
                fileStatusFlag &= ~checkFileStatus.allCustomer; //เคลียร์บิต
                con.Close();
                dr.Close();
            }
  
        }
        //ตรวจสอบและโหลดข้อมูลแสดงใน panel
        private void loadPanelDB()
        {
            try
            {
                //เช็คสถานะฝากมาตร
                con.Open();
                cmd.Connection = con;
                string selectStatusDeposit = "SELECT top 1 รหัสหน่วยงาน FROM depositMeter";
                cmd.CommandText = selectStatusDeposit;
                dr = cmd.ExecuteReader();
                if(dr.HasRows)
                {
                   dr.Read();
                   if (dr[0].ToString() == checkBA)
                   {
                      lblStatusDepositMeter.Text = "R";
                      lblStatusDepositMeter.ForeColor = Color.LightSeaGreen;
                      fileStatusFlag |= checkFileStatus.depositMeter;
                   }
                   else
                   {
                      lblStatusDepositMeter.Text = "Q";
                      lblStatusDepositMeter.ForeColor = Color.PaleVioletRed;
                      fileStatusFlag &= ~checkFileStatus.depositMeter;
                   } 
                }
                else
                {
                    lblStatusDepositMeter.Text = "ฃ";
                    lblStatusDepositMeter.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะหลักประกันสัญญา
                con.Open();
                cmd.Connection = con;
                string selectStatusGaruntee = "SELECT top 1 รหัสหน่วยงาน FROM garuntee";
                cmd.CommandText = selectStatusGaruntee;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusGaruntee.Text = "R";
                        lblStatusGaruntee.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.garuntee;
                    }
                    else
                    {
                        lblStatusGaruntee.Text = "Q";
                        lblStatusGaruntee.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.garuntee;
                    }
//////////////////////////////////////////////////////
                    dr.Close();
                    con.Close();
                    //ตรวจสอบไฟล์เขต
                    con.Open();
                    cmd.Connection = con;
                    string selectlbl = "SELECT top 1 ba.เลขที่สาขา, ba.ORG_OWNER_ID, ba.ชื่อสาขา FROM ba INNER JOIN garuntee ON ba.ORG_OWNER_ID = garuntee.รหัสหน่วยงาน;";
                    cmd.CommandText = selectlbl;
                    dr = cmd.ExecuteReader();
                    dr.Read();
                    string reg = dr[1].ToString();
                    if (isReg(reg))
                    {
                        toolStripLabelUnit.Text = "BA " + dr[0].ToString();
                        lblUnitBAReg = toolStripLabelUnit.Text;
                        checkBAReg = reg;
                        lblStatusGarunteeReg.Text = "R";
                        lblStatusGarunteeReg.ForeColor = Color.LightSeaGreen;
                        comboBoxUnit.Text = "การประปาส่วนภูมิภาค" + dr[2].ToString();
                        string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                        executeInfo(update);
                        fileStatusFlagReg |= checkFileStatusReg.garuntee;
                    }
                    else
                    {
                        lblStatusGarunteeReg.Text = "Q";
                        lblStatusGarunteeReg.ForeColor = Color.PaleVioletRed;
                        fileStatusFlagReg &= ~checkFileStatusReg.garuntee;
                    }                       
///////////////////////////////////////////////////////                   
                }
                else
                {
                    lblStatusGaruntee.Text = "ฃ";
                    lblStatusGaruntee.ForeColor = SystemColors.ControlDarkDark;
                    lblStatusGarunteeReg.Text = "ฃ";
                    lblStatusGarunteeReg.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะรับเงินค่าติดตั้ง
                con.Open();
                cmd.Connection = con;
                string selectStatusInstallCost = "SELECT top 1 รหัสหน่วยงาน FROM installCost";
                cmd.CommandText = selectStatusInstallCost;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusInstallCost.Text = "R";
                        lblStatusInstallCost.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.installCost;
                    }
                    else
                    {
                        lblStatusInstallCost.Text = "Q";
                        lblStatusInstallCost.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.installCost;
                    }
//////////////////////////////////////////////////////
                    dr.Close();
                    con.Close();
                    //ตรวจสอบไฟล์เขต
                    con.Open();
                    cmd.Connection = con;
                    string selectlbl = "SELECT top 1 ba.เลขที่สาขา, ba.ORG_OWNER_ID, ba.ชื่อสาขา FROM ba INNER JOIN installCost ON ba.ORG_OWNER_ID = installCost.รหัสหน่วยงาน;";
                    cmd.CommandText = selectlbl;
                    dr = cmd.ExecuteReader();
                    dr.Read();
                    string reg = dr[1].ToString();
                    if (isReg(reg))
                    {
                        toolStripLabelUnit.Text = "BA " + dr[0].ToString();
                        lblUnitBAReg = toolStripLabelUnit.Text;
                        checkBAReg = reg;
                        lblStatusInstallCostReg.Text = "R";
                        lblStatusInstallCostReg.ForeColor = Color.LightSeaGreen;
                        comboBoxUnit.Text = "การประปาส่วนภูมิภาค" + dr[2].ToString();
                        string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                        executeInfo(update);
                        fileStatusFlagReg |= checkFileStatusReg.installCost;
                    }
                    else
                    {
                        lblStatusInstallCostReg.Text = "Q";
                        lblStatusInstallCostReg.ForeColor = Color.PaleVioletRed;
                        fileStatusFlagReg &= ~checkFileStatusReg.installCost;
                    }
///////////////////////////////////////////////////////
                }
                else
                {
                    lblStatusInstallCost.Text = "ฃ";
                    lblStatusInstallCost.ForeColor = SystemColors.ControlDarkDark;
                    lblStatusInstallCostReg.Text = "ฃ";
                    lblStatusInstallCostReg.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะรับเงินรายได้ค่าน้ำ
                con.Open();
                cmd.Connection = con;
                string selectStatusWaterRevenue = "SELECT top 1 รหัสหน่วยงาน FROM waterRevenue";
                cmd.CommandText = selectStatusWaterRevenue;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusWaterRevenue.Text = "R";
                        lblStatusWaterRevenue.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.waterRevenue;
                    }
                    else
                    {
                        lblStatusWaterRevenue.Text = "Q";
                        lblStatusWaterRevenue.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.waterRevenue;
                    }
//////////////////////////////////////////////////////
                    dr.Close();
                    con.Close();
                    //ตรวจสอบไฟล์เขต
                    con.Open();
                    cmd.Connection = con;
                    string selectlbl = "SELECT top 1 ba.เลขที่สาขา, ba.ORG_OWNER_ID, ba.ชื่อสาขา FROM ba INNER JOIN waterRevenue ON ba.ORG_OWNER_ID = waterRevenue.รหัสหน่วยงาน;";
                    cmd.CommandText = selectlbl;
                    dr = cmd.ExecuteReader();
                    dr.Read();
                    string reg = dr[1].ToString();
                    if (isReg(reg))
                    {
                        toolStripLabelUnit.Text = "BA " + dr[0].ToString();
                        lblUnitBAReg = toolStripLabelUnit.Text;
                        checkBAReg = reg;
                        lblStatusWaterRevenueReg.Text = "R";
                        lblStatusWaterRevenueReg.ForeColor = Color.LightSeaGreen;
                        comboBoxUnit.Text = "การประปาส่วนภูมิภาค" + dr[2].ToString();
                        string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                        executeInfo(update);
                        fileStatusFlagReg |= checkFileStatusReg.waterRevenue;
                    }
                    else
                    {
                        lblStatusWaterRevenueReg.Text = "Q";
                        lblStatusWaterRevenueReg.ForeColor = Color.PaleVioletRed;
                        fileStatusFlagReg &= ~checkFileStatusReg.waterRevenue;
                    }
///////////////////////////////////////////////////////
                }
                else
                {
                    lblStatusWaterRevenue.Text = "ฃ";
                    lblStatusWaterRevenue.ForeColor = SystemColors.ControlDarkDark;
                    lblStatusWaterRevenueReg.Text = "ฃ";
                    lblStatusWaterRevenueReg.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะทะเบียนคุมน้ำท่อธาร
                con.Open();
                cmd.Connection = con;
                string selectStatusPipeMeter = "SELECT top 1 รหัสหน่วยงาน FROM pipeMeter";
                cmd.CommandText = selectStatusPipeMeter;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusPipeMeter.Text = "R";
                        lblStatusPipeMeter.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.pipeMeter;
                    }
                    else
                    {
                        lblStatusPipeMeter.Text = "Q";
                        lblStatusPipeMeter.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.pipeMeter;
                    }
                }
                else
                {
                    lblStatusPipeMeter.Text = "ฃ";
                    lblStatusPipeMeter.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะตั้งหนี้ประจำเดือน
                con.Open();
                cmd.Connection = con;
                string selectStatusDebtMonth = "SELECT top 1 รหัสหน่วยงาน FROM debtMonth";
                cmd.CommandText = selectStatusDebtMonth;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusDebtMonth.Text = "R";
                        lblStatusDebtMonth.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.debtMonth;
                    }
                    else
                    {
                        lblStatusDebtMonth.Text = "Q";
                        lblStatusDebtMonth.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.debtMonth;
                    }
                }
                else
                {
                    lblStatusDebtMonth.Text = "ฃ";
                    lblStatusDebtMonth.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะตั้งหนี้ผู้ใช้น้ำ
                con.Open();
                cmd.Connection = con;
                string selectStatusDebtCustomer = "SELECT top 1 รหัสหน่วยงาน FROM debtCustomer";
                cmd.CommandText = selectStatusDebtCustomer;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusDebtCustomer.Text = "R";
                        lblStatusDebtCustomer.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.debtCustomer;
                    }
                    else
                    {
                        lblStatusDebtCustomer.Text = "Q";
                        lblStatusDebtCustomer.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.debtCustomer;
                    }
                }
                else
                {
                    lblStatusDebtCustomer.Text = "ฃ";
                    lblStatusDebtCustomer.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะรายได้อื่นๆเกี่ยวกับผู้ใช้น้ำ
                con.Open();
                cmd.Connection = con;
                string selectStatusOtherRevenueCustomer = "SELECT top 1 รหัสหน่วยงาน FROM otherRevenueCustomer";
                cmd.CommandText = selectStatusOtherRevenueCustomer;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusOtherRevenueCustomer.Text = "R";
                        lblStatusOtherRevenueCustomer.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.otherRevenueCustomer;
                    }
                    else
                    {
                        lblStatusOtherRevenueCustomer.Text = "Q";
                        lblStatusOtherRevenueCustomer.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.otherRevenueCustomer;
                    }
                }
                else
                {
                    lblStatusOtherRevenueCustomer.Text = "ฃ";
                    lblStatusOtherRevenueCustomer.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะรายได้อื่นๆไม่เกี่ยวกับผู้ใช้น้ำ
                con.Open();
                cmd.Connection = con;
                string selectStatusOtherRevenueNonCustomer = "SELECT top 1 รหัสหน่วยงาน FROM otherRevenueNonCustomer";
                cmd.CommandText = selectStatusOtherRevenueNonCustomer;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusOtherRevenueNonCustomer.Text = "R";
                        lblStatusOtherRevenueNonCustomer.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.otherRevenueNonCustomer;
                    }
                    else
                    {
                        lblStatusOtherRevenueNonCustomer.Text = "Q";
                        lblStatusOtherRevenueNonCustomer.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.otherRevenueNonCustomer;
                    }
//////////////////////////////////////////////////////
                    dr.Close();
                    con.Close();
                    //ตรวจสอบไฟล์เขต
                    con.Open();
                    cmd.Connection = con;
                    string selectlbl = "SELECT top 1 ba.เลขที่สาขา, ba.ORG_OWNER_ID, ba.ชื่อสาขา FROM ba INNER JOIN otherRevenueNonCustomer ON ba.ORG_OWNER_ID = otherRevenueNonCustomer.รหัสหน่วยงาน;";
                    cmd.CommandText = selectlbl;
                    dr = cmd.ExecuteReader();
                    dr.Read();
                    string reg = dr[1].ToString();
                    if (isReg(reg))
                    {
                        toolStripLabelUnit.Text = "BA " + dr[0].ToString();
                        lblUnitBAReg = toolStripLabelUnit.Text;
                        checkBAReg = reg;
                        lblStatusOtherRevenueNonCustomerReg.Text = "R";
                        lblStatusOtherRevenueNonCustomerReg.ForeColor = Color.LightSeaGreen;
                        comboBoxUnit.Text = "การประปาส่วนภูมิภาค" + dr[2].ToString();
                        string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                        executeInfo(update);
                        fileStatusFlagReg |= checkFileStatusReg.otherRevenueNonCustomer;
                    }
                    else
                    {
                        lblStatusOtherRevenueNonCustomerReg.Text = "Q";
                        lblStatusOtherRevenueNonCustomerReg.ForeColor = Color.PaleVioletRed;
                        fileStatusFlagReg &= ~checkFileStatusReg.otherRevenueNonCustomer;
                    }
///////////////////////////////////////////////////////
                }
                else
                {
                    lblStatusOtherRevenueNonCustomer.Text = "ฃ";
                    lblStatusOtherRevenueNonCustomer.ForeColor = SystemColors.ControlDarkDark;
                    lblStatusOtherRevenueNonCustomerReg.Text = "ฃ";
                    lblStatusOtherRevenueNonCustomerReg.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะหนี้ค้างชำระ
                con.Open();
                cmd.Connection = con;
                string selectStatusDebtCurrent = "SELECT top 1 รหัสหน่วยงาน FROM debtCurrent";
                cmd.CommandText = selectStatusDebtCurrent;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusDebtCurrent.Text = "R";
                        lblStatusDebtCurrent.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.debtCurrent;
                    }
                    else
                    {
                        lblStatusDebtCurrent.Text = "Q";
                        lblStatusDebtCurrent.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.debtCurrent;
                    }
                }
                else
                {
                    lblStatusDebtCurrent.Text = "ฃ";
                    lblStatusDebtCurrent.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะประวัติการเปลี่ยนสถานะมาตร
                con.Open();
                cmd.Connection = con;
                string selectStatusMeterHistory = "SELECT top 1 รหัสหน่วยงาน FROM meterHistory";
                cmd.CommandText = selectStatusMeterHistory;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusMeterHistory.Text = "R";
                        lblStatusMeterHistory.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.meterHistory;
                    }
                    else
                    {
                        lblStatusMeterHistory.Text = "Q";
                        lblStatusMeterHistory.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.meterHistory;
                    }
                }
                else
                {
                    lblStatusMeterHistory.Text = "ฃ";
                    lblStatusMeterHistory.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะมาตรตาย
                con.Open();
                cmd.Connection = con;
                string selectStatusMeterAbnormal = "SELECT top 1 รหัสหน่วยงาน FROM meterAbnormal";
                cmd.CommandText = selectStatusMeterAbnormal;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusMeterAbnormal.Text = "R";
                        lblStatusMeterAbnormal.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.meterAbnormal;
                    }
                    else
                    {
                        lblStatusMeterAbnormal.Text = "Q";
                        lblStatusMeterAbnormal.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.meterAbnormal;
                    }
                }
                else
                {
                    lblStatusMeterAbnormal.Text = "ฃ";
                    lblStatusMeterAbnormal.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                ///เช็คสถานะยกเลิกใบเสร็จ
                con.Open();
                cmd.Connection = con;
                string selectStatusCancelReceipt = "SELECT top 1 รหัสหน่วยงาน FROM Repaid";
                cmd.CommandText = selectStatusCancelReceipt;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() == checkBA)
                    {
                        lblStatusCancelReceipt.Text = "R";
                        lblStatusCancelReceipt.ForeColor = Color.LightSeaGreen;
                        fileStatusFlag |= checkFileStatus.cancelReceipt;
                    }
                    else
                    {
                        lblStatusCancelReceipt.Text = "Q";
                        lblStatusCancelReceipt.ForeColor = Color.PaleVioletRed;
                        fileStatusFlag &= ~checkFileStatus.cancelReceipt;
                    }
                    //////////////////////////////////////////////////////
                    dr.Close();
                    con.Close();
                    //ตรวจสอบไฟล์เขต
                    con.Open();
                    cmd.Connection = con;
                    string selectlbl = "SELECT top 1 ba.เลขที่สาขา, ba.ORG_OWNER_ID, ba.ชื่อสาขา FROM ba INNER JOIN Repaid ON ba.ORG_OWNER_ID = Repaid.รหัสหน่วยงาน;";
                    cmd.CommandText = selectlbl;
                    dr = cmd.ExecuteReader();
                    dr.Read();
                    string reg = dr[1].ToString();
                    if (isReg(reg))
                    {
                        toolStripLabelUnit.Text = "BA " + dr[0].ToString();
                        lblUnitBAReg = toolStripLabelUnit.Text;
                        checkBAReg = reg;
                        lblStatusCancelReceipt.Text = "R";
                        lblStatusCancelReceipt.ForeColor = Color.LightSeaGreen;
                        comboBoxUnit.Text = "การประปาส่วนภูมิภาค" + dr[2].ToString();
                        string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                        executeInfo(update);
                        fileStatusFlagReg |= checkFileStatusReg.cancelReceipt;
                    }
                    else
                    {
                        lblStatusCancelReceiptReg.Text = "Q";
                        lblStatusCancelReceiptReg.ForeColor = Color.PaleVioletRed;
                        fileStatusFlagReg &= ~checkFileStatusReg.cancelReceipt;
                    }
                    ///////////////////////////////////////////////////////                   
                }
                else
                {
                    lblStatusCancelReceipt.Text = "ฃ";
                    lblStatusCancelReceipt.ForeColor = SystemColors.ControlDarkDark;
                    lblStatusCancelReceiptReg.Text = "ฃ";
                    lblStatusCancelReceiptReg.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

                //เช็คสถานะหนี้อยู่ระหว่างดำเนินคดี
                con.Open();
                cmd.Connection = con;
                string selectStatusDebtLitigate = "SELECT top 1 รหัสหน่วยงาน FROM debtLitigate";
                cmd.CommandText = selectStatusDebtLitigate;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dr.Read();
                    if (dr[0].ToString() != "")
                    {
                        lblStatusLitigate.Text = "R";
                        lblStatusLitigate.ForeColor = Color.LightSeaGreen;
                        fileStatusFlagReg |= checkFileStatusReg.debtLitigate;
                    }
                    else
                    {
                        lblStatusLitigate.Text = "Q";
                        lblStatusLitigate.ForeColor = Color.PaleVioletRed;
                        fileStatusFlagReg &= ~checkFileStatusReg.debtLitigate;
                    }
                }
                else
                {
                    lblStatusLitigate.Text = "ฃ";
                    lblStatusLitigate.ForeColor = SystemColors.ControlDarkDark;
                }
                con.Close();
                dr.Close();

            }
            catch
            {
                con.Close();
                dr.Close();
            }
        }

        //เมทตอทเชื่อมต่อและอัพเดทฐานข้อมูล Info
        private void executeInfo (string exe)
        {
            try
            {
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = exe;
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception)
            {
                con.Close(); 
            }

        }
        
        //โหลดข้อมูลจากฐานข้อมูล Info
        private void frmMain_Load(object sender, EventArgs e)
        {
            con.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + homePath + "\\userinfo\\info.mdb; Jet OLEDB:Database Password = pwacis2561;";
            cmd.Connection = con;
            loadInfo();
            loadCustomerDB();
            loadPanelDB();
            dtpReviewDate.MinDate = dtpAuditDate.Value;
            dtpPeriodBegin.MaxDate = dtpPeriodEnd.Value;
            dtpPeriodEnd.MinDate = dtpPeriodBegin.Value;
            
            setButton();
            setButtonReg();
        }
        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
           
            if(checkLoad == 2)
            {
                e.Cancel = true;
            }
            else
            {
                int checkReg = 0; //ตรวจสอบสถานะประปา = 0 เขต = 1
                if (checkBoxUnitName.Checked)
                {
                    checkReg = 1;

                    if (checkBoxReviewDate.Checked)
                    {
                        dtpReviewDate.Enabled = false;
                        checkBoxReview = "null";
                        string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate =  " + checkBoxReview + " , checkReg =  " + checkReg + "  where ID = 1";
                        executeInfo(update);
                    }
                    else
                    {
                        dtpReviewDate.Enabled = true;
                        checkBoxReview = "'" + dtpReviewDate.Text + "'";
                        string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate =  " + checkBoxReview + " , checkReg =  " + checkReg + " where ID = 1";
                        executeInfo(update);
                    }

                }
                else
                {
                    if (checkBoxReviewDate.Checked)
                    {
                        dtpReviewDate.Enabled = false;
                        checkBoxReview = "null";
                        string update = "UPDATE info SET unitName = '" + txtUnitName.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate =  " + checkBoxReview + " , checkReg =  " + checkReg + "  where ID = 1";
                        executeInfo(update);
                    }
                    else
                    {
                        dtpReviewDate.Enabled = true;
                        checkBoxReview = "'" + dtpReviewDate.Text + "'";
                        string update = "UPDATE info SET unitName = '" + txtUnitName.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate =  " + checkBoxReview + " , checkReg =  " + checkReg + " where ID = 1";
                        executeInfo(update);
                    }
                }
               
      
            }
            checkLoad = 0;
        }
        private void txtUnitName_Leave(object sender, EventArgs e)
        {
            //unitName = '" + txtUnitName.Text + "',
            string update = "UPDATE info SET unitName = '" + txtUnitName.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
            executeInfo(update);            
        }     
        private void txtAuditName_Leave(object sender, EventArgs e)
        {
            //unitName = '" + txtUnitName.Text + "',
            string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
            executeInfo(update);
        }
        private void txtReviewName_Leave(object sender, EventArgs e)
        {
            //unitName = '" + txtUnitName.Text + "',
            string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
            executeInfo(update);
        }
        private void dtpPeriodBegin_Leave(object sender, EventArgs e)
        {
            //unitName = '" + txtUnitName.Text + "',
            string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
            executeInfo(update);
            dtpPeriodEnd.MinDate = dtpPeriodBegin.Value;
        }
        private void dtpPeriodEnd_Leave(object sender, EventArgs e)
        {
            string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
            executeInfo(update);
            dtpPeriodBegin.MaxDate = dtpPeriodEnd.Value;
        }
        private void dtpAuditDate_Leave(object sender, EventArgs e)
        {
            if (dtpReviewDate.Value < dtpAuditDate.Value)
            {
                dtpReviewDate.Value = dtpAuditDate.Value;
                string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";            
                executeInfo(update);           
            }
            else
            {
                string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            dtpReviewDate.MinDate = dtpAuditDate.Value;
        }
        private void dtpReviewDate_Leave(object sender, EventArgs e)
        {
            string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
            executeInfo(update);
        }

        //Status Bar 
        private void frmMain_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "";
            toolStripStatusLabelBottom.ForeColor = Color.FromArgb(0 ,0 ,0);
            
        }
        private void txtUnitName_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "กรอกชื่อหน่วยรับตรวจ เช่น การประปาส่วนภูมิภาคสาขา...";
        }
        private void txtAuditName_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "กรอกชื่อผู้ตรวจสอบ";
        }
        private void txtReviewName_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "กรอกชื่อผู้สอบทาน";
        }
        private void dtpPeriodBegin_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "ระบุวันที่งวดการตรวจสอบเริ่มต้น";
        }
        private void dtpPeriodEnd_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "ระบุวันที่งวดการตรวจสอบสิ้นสุด";
        }
        private void dtpAuditDate_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "ระบุวันที่ตรวจสอบ";
        }
        private void dtpReviewDate_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "ระบุวันที่สอบทาน";
        }
        private void btnOpenWeb_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.Text = "เปิดเว็บไซต์เพื่อโหลดฐานข้อมูล";
        }
        private void checkBoxReviewDate_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabelBottom.ForeColor = Color.FromArgb(255, 0, 0);
            toolStripStatusLabelBottom.Text = "เลือกเพื่อไม่ระบุวันที่ผู้สอบทาน";
        }
        
        //เช็ควันที่ผู้สอบทาน
        private void checkBoxReviewDate_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxReviewDate.Checked)
            {
                dtpReviewDate.Enabled = false;
                checkBoxReview = "null";
                //unitName = '" + txtUnitName.Text + "',
                string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate =  " + checkBoxReview + "  where ID = 1";
                executeInfo(update);
            }
            else
            {
                dtpReviewDate.Enabled = true;
                checkBoxReview = "'" + dtpReviewDate.Text + "'";
                //unitName = '" + txtUnitName.Text + "',
                string update = "UPDATE info SET auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate =  " + checkBoxReview + "  where ID = 1";
                executeInfo(update);
            }

        }

        //เช็คว่าเป็นเขตหรือประปา
        private void checkBoxUnitName_CheckedChanged(object sender, EventArgs e)
        {
            
            if (checkBoxUnitName.Checked)
            {
               // panelMain.Hide();
               // panelWorkingPaper.Hide();
               // txtUnitName.Hide();
                panelMainReg.Show();
                panelWorkingPaperReg.Show();
                comboBoxUnit.Show();
                toolStripLabelUnit.Text = lblUnitBAReg;
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else
            {
               // panelMain.Show();
               // panelWorkingPaper.Show();
               // txtUnitName.Show();
                panelMainReg.Hide();
                panelWorkingPaperReg.Hide();
                comboBoxUnit.Hide();
                toolStripLabelUnit.Text = lblUnitBA;
                string update = "UPDATE info SET unitName = '" + txtUnitName.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
        }

        //ปุ่ม Exit
        private void buttonExit_MouseHover(object sender, EventArgs e)
        {
            buttonExit.BackgroundImage = CISA.Properties.Resources.exit_button_md_red;
            toolStripStatusLabelBottom.ForeColor = Color.FromArgb(255, 0, 0);
            toolStripStatusLabelBottom.Text = "ออกจากโปรแกรม";
        }
        private void buttonExit_MouseLeave(object sender, EventArgs e)
        {
            buttonExit.BackgroundImage = CISA.Properties.Resources.exit_button_md;
        }
        private void buttonExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        
        //ปุ่ม Upload
        private void toolStripButtonUpload_MouseHover(object sender, EventArgs e)
        {
            toolStripButtonUpload.BackgroundImage = CISA.Properties.Resources.upload_red;
            toolStripStatusLabelBottom.ForeColor = Color.FromArgb(255, 0, 0);
            toolStripStatusLabelBottom.Text = "อัพโหลดไฟล์ข้อมูลที่โหลดจากเว็บ Support เท่านั้น ไม่สามารถอัพโหลดไฟล์อื่นได้";
        }
        private void toolStripButtonUpload_MouseLeave(object sender, EventArgs e)
        {
            toolStripButtonUpload.BackgroundImage = CISA.Properties.Resources.upload;
        }

        //อัพโหลดไฟล์ข้อมูลผู้ใช้น้ำทั้งหมด       
        private void toolStripButtonUpload_Click(object sender, EventArgs e)
        {
            ope = new OpenFileDialog();
            ope.Multiselect = true;
            ope.Title = "เลือกไฟล์ *.xlsx ที่โหลดมาจากเว็บไซต์ที่ได้จัดทำไว้เท่านั้น";

            ope.Filter = "xlsx Files(*.xlsx)|*.xlsx|All Files(*.xlsx)|*.xlsx";
            
            if (ope.ShowDialog() == DialogResult.OK)
            {
                filesOpen = ope.FileNames;

                //ปิดการทำงานหน้าเมนู
                progressBar.Visible = true;
                toolStripButtonUpload.Enabled = false;
                txtUnitName.Enabled = false;
                txtAuditName.Enabled = false;
                txtReviewName.Enabled = false;
                dtpPeriodBegin.Enabled = false;
                dtpPeriodEnd.Enabled = false;
                dtpAuditDate.Enabled = false;
                dtpReviewDate.Enabled = false;
                checkBoxReviewDate.Enabled = false;
                buttonExit.Enabled = false;
                checkBoxUnitName.Enabled = false;
                panelWorkingPaper.Enabled = false;
                panelWorkingPaperReg.Enabled = false;
                
                checkLoad = 2;
             
                backgroundWorker.RunWorkerAsync();

            }

        }

        private void frmMain_DragDrop(object sender, DragEventArgs e)
        {
            filesOpen = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            progressBar.Visible = true;
            toolStripButtonUpload.Enabled = false;
            txtUnitName.Enabled = false;
            txtAuditName.Enabled = false;
            txtReviewName.Enabled = false;
            dtpPeriodBegin.Enabled = false;
            dtpPeriodEnd.Enabled = false;
            dtpAuditDate.Enabled = false;
            dtpReviewDate.Enabled = false;
            checkBoxReviewDate.Enabled = false;
            buttonExit.Enabled = false;
            checkBoxUnitName.Enabled = false;
            panelWorkingPaper.Enabled = false;
            panelWorkingPaperReg.Enabled = false;
            
            checkLoad = 2;

            backgroundWorker.RunWorkerAsync();

        }

        private void frmMain_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
            checkLoad = 0;
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            
            foreach (string filePaths in filesOpen)
            {

                filePath = filePaths;
                if(Path.GetExtension(filePath) == ".xlsx")
                {
                    try
                    {
                        FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        DataSet result = excelReader.AsDataSet();

                        if (result.Tables.Count == 0)
                        {
                            MessageBox.Show("ไม่พบข้อมูลในไฟล์\n"+ filePath, "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            checkLoad = 0;
                        }
                        else
                        {
                            if (result.Tables[0].TableName == "ผู้ใช้น้ำทั้งหมด")
                            {
                                string drop = "drop table customer";
                                executeInfo(drop);
                                
                                string upload = @"select * into customer from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[ผู้ใช้น้ำทั้งหมด$] s;";
                                executeInfo(upload);
                                /*
                                string convert = "alter table customer alter column วันที่ให้สิทธิส่วนลด date";
                                executeInfo(convert);
                                convert = "alter table customer alter column วันที่ครบกำหนดสิทธิส่วนลด Date";
                                executeInfo(convert);
                                */
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "ฝากมาตรครบกำหนด")
                            {
                                string drop = "drop table depositMeter";
                                executeInfo(drop);
                                string upload = @"select * into depositMeter from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[ฝากมาตรครบกำหนด$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "หลักประกันสัญญา")
                            {
                                string drop = "drop table garuntee";
                                executeInfo(drop);
                                string upload = @"select * into garuntee from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[หลักประกันสัญญา$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "รับเงินค่าติดตั้ง")
                            {
                                string drop = "drop table installCost";
                                executeInfo(drop);
                                string upload = @"select * into installCost from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[รับเงินค่าติดตั้ง$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "รับเงินรายได้ค่าน้ำ")
                            {
                                string drop = "drop table waterRevenue";
                                executeInfo(drop);
                                string upload = @"select * into waterRevenue from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[รับเงินรายได้ค่าน้ำ$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "ทะเบียนคุมน้ำท่อธาร")
                            {
                                string drop = "drop table pipeMeter";
                                executeInfo(drop);
                                string upload = @"select * into pipeMeter from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[ทะเบียนคุมน้ำท่อธาร$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "ตั้งหนี้ประจำเดือน")
                            {
                                string drop = "drop table debtMonth";
                                executeInfo(drop);
                                string upload = @"select * into debtMonth from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[ตั้งหนี้ประจำเดือน$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "ตั้งหนี้ผู้ใช้น้ำ")
                            {
                                string drop = "drop table debtCustomer";
                                executeInfo(drop);
                                string upload = @"select * into debtCustomer from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[ตั้งหนี้ผู้ใช้น้ำ$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "รายได้อื่นๆ ที่เกี่ยวกับผู้ใช้น")
                            {
                                string drop = "drop table otherRevenueCustomer";
                                executeInfo(drop);
                                string upload = @"select * into otherRevenueCustomer from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[รายได้อื่นๆ ที่เกี่ยวกับผู้ใช้น$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "รายได้อื่นๆ ที่ไม่เกี่ยวกับผู้ใ")
                            {
                                string drop = "drop table otherRevenueNonCustomer";
                                executeInfo(drop);
                                string upload = @"select * into otherRevenueNonCustomer from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[รายได้อื่นๆ ที่ไม่เกี่ยวกับผู้ใ$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "หนี้ค้างชำระ")
                            {
                                string drop = "drop table debtCurrent";
                                executeInfo(drop);
                                string upload = @"select * into debtCurrent from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[หนี้ค้างชำระ$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "ประวัติการเปลี่ยนสถานะมาตร")
                            {
                                string drop = "drop table meterHistory";
                                executeInfo(drop);
                                string upload = @"select * into meterHistory from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[ประวัติการเปลี่ยนสถานะมาตร$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "มาตรตาย")
                            {
                                string drop = "drop table meterAbnormal";
                                executeInfo(drop);
                                string upload = @"select * into meterAbnormal from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[มาตรตาย$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else if (result.Tables[0].TableName == "ยกเลิกใบเสร็จรับเงินค่าน้ำ")
                            {
                                string drop = "drop table Repaid";
                                executeInfo(drop);
                                string upload = @"select * into Repaid from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[ยกเลิกใบเสร็จรับเงินค่าน้ำ$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            //รอจัดทำข้อมูลจาก CIS Suport เพื่อเปลี่ยนชื่อTableตามExcel
                            else if (result.Tables[0].TableName == "หนี้ดำเนินคดี")
                            {
                                string drop = "drop table debtLitigate";
                                executeInfo(drop);
                                string upload = @"select * into debtLitigate from [Excel 8.0;HDR=YES;DATABASE=" + filePath + "].[หนี้ดำเนินคดี$] s;";
                                executeInfo(upload);
                                checkLoad = 1;
                            }
                            else
                            {
                                MessageBox.Show("เลือกไฟล์ข้อมูลไม่ถูกต้อง\n" + filePath + " \nกรุณาเลือกไฟล์ใหม่", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                checkLoad = 0;
                            }
                        }

                    }
                    catch
                    {
                        MessageBox.Show("กรุณาปิดไฟล์ก่อนทำการอัพโหลดข้อมูลหรือพบว่ามีการเลือกไฟล์ไม่ถูกต้อง", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        checkLoad = 0;
                    }
                }

            }

        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //เปิดการใช้งานหน้าเมนู
            progressBar.Visible = false;
            toolStripButtonUpload.Enabled = true;
            txtUnitName.Enabled = true;
            txtAuditName.Enabled = true;
            txtReviewName.Enabled = true;
            dtpPeriodBegin.Enabled = true;
            dtpPeriodEnd.Enabled = true;
            dtpAuditDate.Enabled = true;
            dtpReviewDate.Enabled = true;
            checkBoxReviewDate.Enabled = true;
            buttonExit.Enabled = true;
            panelWorkingPaper.Enabled = true;
            panelWorkingPaperReg.Enabled = true;
            checkBoxUnitName.Enabled = true;
            txtAuditName.Focus(); //โฟกัสเพื่อบันทึกค่า
            try
            {
                con.Open();
                cmd.Connection = con;
                string select = "SELECT top 1 ba.ชื่อสาขา FROM ba INNER JOIN customer ON ba.ORG_OWNER_ID = customer.รหัสหน่วยงาน;";
                cmd.CommandText = select;
                dr = cmd.ExecuteReader();
                dr.Read();
                txtUnitName.Text = "การประปาส่วนภูมิภาค" + dr[0].ToString();
                con.Close();
                if(checkLoad == 1)
                {
                    MessageBox.Show("อัพโหลดข้อมูลเสร็จสิ้น", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    checkLoad = 0;
                }
                loadCustomerDB();
                loadPanelDB();
                setButton();
                setButtonReg();
                checkLoad = 0;
            }
            catch (Exception)
            {
                MessageBox.Show("ไม่พบข้อมูลในไฟล์หรือข้อมูลไม่ถูกต้อง","พบข้อผิดพลาด",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                checkLoad = 0;
            }
        
        }


        //ปุ่มเรียกใช้งานมาโคร excel
        private void btnRegisterCustomer_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\1wp_newcustomer_v1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCalWater_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\2wp_cal_bill_V1a.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnDiscount_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\3wp_discount_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnMeterAbnormal_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\4wp_abnormalMeter_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnAdjustDebt_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\5wp_adjustDebt_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnRandomMeter_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\6wp_randommeter_V1c.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
        }

        private void btnPipe_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\7wp_pipe_v1a.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnOtherRevenueCustomer_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\8wp_Other_Customer_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnOtherRevenueNonCustomer_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\9wp_Other_nonCustomer_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnDepositMeter_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\10wp_Deposit_Meter_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnReconcileMoney_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\11wp_reconcileMoney_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnConfirmBalance_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\12wp_confirmBalance_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnDebtCurrent_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\13wp_debtCurrent_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnInstallCost_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\14wp_installCost_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnGaruntee_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\15wp_bank_guarantee_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnCancelReceipt_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\16wp_cancel_receipt_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //ปุ่มเรียกใช้งานมาโคร excel เขต
        private void BtnOtherRevenueNonCustomerReg_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\9wp_Other_nonCustomer_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnInstallCostReg_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\14wp_installCost_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnGarunteeReg_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\15wp_bank_guarantee_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCancelReceiptReg_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\16wp_cancel_receipt_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnLitigateAll_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(homePath + "\\wp_macro\\17wp_debtLitigate_V1.xlsm");
            }
            catch
            {
                MessageBox.Show("ไม่พบไฟล์โปรแกรมข้อมูล", "พบข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        //ComboboxUnit ของเขต
        private void comboBoxUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 1")
            {
                toolStripLabelUnit.Text = "BA 1101";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "1101";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 2")
            {
                toolStripLabelUnit.Text = "BA 1147";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "1147";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 3")
            {
                toolStripLabelUnit.Text = "BA 1178";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "178";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 4")
            {
                toolStripLabelUnit.Text = "BA 1203";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "203";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 5")
            {
                toolStripLabelUnit.Text = "BA 1225";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "1225";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 6")
            {
                toolStripLabelUnit.Text = "BA 1059";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "1059";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 7")
            {
                toolStripLabelUnit.Text = "BA 1078";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "1078";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 8")
            {
                toolStripLabelUnit.Text = "BA 1124";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "1124";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 9")
            {
                toolStripLabelUnit.Text = "BA 1003";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "10003";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
            else if (comboBoxUnit.Text == "การประปาส่วนภูมิภาคเขต 10")
            {
                toolStripLabelUnit.Text = "BA 1031";
                lblUnitBAReg = toolStripLabelUnit.Text;
                checkBAReg = "1031";
                string update = "UPDATE info SET unitName = '" + comboBoxUnit.Text + "', auditName = '" + txtAuditName.Text + "', reviewName = '" + txtReviewName.Text + "', periodBegin = '" + dtpPeriodBegin.Text + "', periodEnd = '" + dtpPeriodEnd.Text + "', auditDate = '" + dtpAuditDate.Text + "', reviewDate = '" + dtpReviewDate.Text + "'  where ID = 1";
                executeInfo(update);
            }
        }

        bool isReg (string ba)
        {
            
            if (ba == "1101"||ba == "1147"||ba == "178" || ba == "203" || ba == "1225" || ba == "1059" || ba == "1078" || ba == "1124" || ba == "10003" || ba == "1031")
            {
                return true;
            }
            return false;                  
        }

        private void btnOpenWeb_Click(object sender, EventArgs e)
        {
            Process.Start("http://servcis.pwa.co.th:8001/CISSupport/index.jsp?menuId=MzQ=");
        }


    }
}
