using Excel;
using ShipWebServiceClient.ShipServiceWebReference;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Protocols;
using System.Windows.Forms;

namespace FedExShipping
{
    public partial class FedExForm : Form
    {
        bool _IsSaved = false, _IsVoid = false;
        string Remarks = "";
        DataTable dtLineItemData = null;
        DataTable table;
        string ExecStatusinfo = string.Empty;
        string TrackingId = string.Empty;
        bool _IsSucceded = false;
        string _FailReason = string.Empty;
        ProcessShipmentRequest childRequest = null;
        List<SecondaryLineItems> listLineItemSecondValues = null;

        public FedExForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            lblUserName.Text = "Welcome Andy Pan";
            lblDateTime.Text = DateTime.Now.DayOfWeek.ToString() + " , " + DateTime.Now.ToShortDateString();
            lblInfo.Text = "Browse File to Submit & Select Response File Path";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv";
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            lblInfo.Text = string.Empty;
            txtResponse.Text = string.Empty;
            try
            {
                DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
                if (result == DialogResult.OK) // Test result.
                {
                    txtFilePath.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception Ex)
            {
                lblInfo.Text = "Exception while Browsing file";
                Remarks += $"BtnBrowse - {Ex}";
            }
        }

        private void btnLineItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult LineItemResult = openFileDialog1.ShowDialog(); // Show the dialog.
                if (LineItemResult == DialogResult.OK) // Test result.
                {
                    txtLineItem.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception Ex)
            {
                lblInfo.Text = "Exception while Browsing Line Item";
                Remarks += $"BtnBrowse - {Ex}";
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                table = new DataTable();
                if (txtFilePath.Text.Trim() == "")
                {
                    MessageBox.Show("OOOPS!!! First Select Data File to Submit !");
                    return;
                }

                if (txtLineItem.Text.Trim() == "")
                {
                    MessageBox.Show("OOOPS!!! First Select Line Item file Path !");
                    return;
                }

                if (txtLblPath.Text.Trim() == "")
                {
                    MessageBox.Show("OOOPS!!! First Select Labels Path !");
                    return;
                }
                ExecStatusinfo = "\n\n";
                MessageBox.Show("Please close selected files while processing....!");
                btnSubmit.Enabled = false;
                DialogResult res = MessageBox.Show("Are you Sure you want to Submit ?", "Ready To Submit!", MessageBoxButtons.YesNo);
                if (res.Equals(DialogResult.Yes))
                {
                    //btnSubmit.Enabled = false;
                    DataTable dtPilotExcelData = ReadExcelData(txtFilePath.Text);
                    dtLineItemData = ReadLineItemData(txtLineItem.Text);
                    if (dtPilotExcelData == null || dtLineItemData == null)
                        return;

                    lblInfo.Text = "Please be patient while processing your request !!";
                    bool _IsFirstRow = true;

                    foreach (DataRow dr in dtPilotExcelData.Rows)
                    {
                        if (_IsFirstRow)
                        {
                            _IsFirstRow = false;
                            continue;
                        }
                        TrackingId = string.Empty;

                        if (dr.ItemArray[22].ToString().Trim() == string.Empty)
                            continue;

                        if (dr.ItemArray[23].ToString().Trim() == "FedEx Home Delivery" ||
                            dr.ItemArray[23].ToString().Trim().ToLower() == "fedex ground")
                        {
                            try
                            {
                                _IsSucceded = false;
                                _FailReason = string.Empty;
                                FedExOperations(dr);
                                if (_IsSucceded)
                                {
                                    ExecStatusinfo += $"PO No# {dr.ItemArray[0].ToString()}     Order No# {dr.ItemArray[1].ToString()}  " +
                                    $"TrackingId# {TrackingId}     Status # Success \r\n";
                                }
                                else
                                {
                                    ExecStatusinfo += $"PO No# {dr.ItemArray[0].ToString()}     Order No# {dr.ItemArray[1].ToString()}  " +
                                   $"TrackingId# {TrackingId}     Status # Failed   Reason#  {_FailReason} \r\n";
                                }
                            }
                            catch (Exception Ex)
                            {
                                ExecStatusinfo += $"PO No# {dr.ItemArray[0].ToString()}     Order No# {dr.ItemArray[1].ToString()}   " +
                                    $"TrackingId# {TrackingId}   Status # Failed Exception ::{Ex.Message} \r\n";
                            }
                            finally
                            {
                                ExecStatusinfo += $"\r\n---------------------------------------------------------------------" +
                                    $"-------------------------\r\n";
                            }
                        }
                        else
                        {
                            _IsSaved = false;
                            _IsVoid = false;
                            Remarks = "Invalid Carrier";
                            //TableOperations(dr.ItemArray[0].ToString(), dr.ItemArray[1].ToString(), dr.ItemArray[25].ToString());
                        }
                    }
                    txtResponse.Text = ExecStatusinfo;
                    lblInfo.Text = "Request Submitted Succesfully.";
                }
            }
            catch (Exception ex)
            {
                lblInfo.Text = "Exception while submitting!";
                Remarks += $"BtnSubmit - {ex}";
            }
            finally
            {
                //ExportToExcel();
                txtFilePath.Text = "";

                //txtResponsePath.Text = "";
                btnBrowse.Enabled = true;
                btnSubmit.Enabled = true;
            }
        }

        public DataTable ReadExcelData(string fileName)
        {
            IExcelDataReader excelReader = null;
            FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            try
            {
                HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(fileName);
                webRequest.ContentType = "application/ms-excel";
                WebResponse objResponse = webRequest.GetResponse();
                Stream excelStream = objResponse.GetResponseStream();

                // your code snippet here
                MemoryStream ms = new MemoryStream();
                byte[] byteBucket = new byte[100];
                int currentByteRead = excelStream.Read(byteBucket, 0, 100);
                while (currentByteRead > 0)
                {
                    ms.Write(byteBucket, 0, currentByteRead);
                    currentByteRead = excelStream.Read(byteBucket, 0, 100);
                }
                ms.Position = 0;

                // using ExcelDataReader to read stream
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(ms);
                // to display the content or other work...
            }
            catch (Exception ex)
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            //excelReader.IsFirstRowAsColumnNames = true;
            DataSet result2 = excelReader.AsDataSet();
            excelReader.Close();
            return result2.Tables[0];
        }

        public DataTable ReadLineItemData(string fileName)
        {
            FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = null;
            try
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            catch
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            //excelReader.IsFirstRowAsColumnNames = true;
            DataSet result2 = excelReader.AsDataSet();
            excelReader.Close();
            return result2.Tables[0];
        }

        private void FedExOperations(DataRow dr)
        {
            // Set this to true to process a COD shipment and print a COD return Label
            bool isCodShipment = false;
            listLineItemSecondValues = null;
            ProcessShipmentRequest request = CreateShipmentRequest(isCodShipment, dr);
            string MasterTrackingNo = string.Empty;
            //
            ShipWebServiceClient.ShipServiceWebReference.ShipService service = new ShipWebServiceClient.ShipServiceWebReference.ShipService();
            if (usePropertyFile())
            {
                service.Url = getProperty("endpoint");
            }
            //
            try
            {
                for (int i = 1; i <= Convert.ToInt16(dr.ItemArray[20].ToString().Trim()); i++)
                {
                    //childRequest = request;
                    //request.RequestedShipment.MasterTrackingId = new ShipWebServiceClient.ShipServiceWebReference.TrackingId();
                    //request.RequestedShipment.MasterTrackingId.TrackingNumber = "0";
                    // Call the ship web service passing in a ProcessShipmentRequest and returning a 
                    //request.RequestedShipment.PackageCount = "2";
                    if (i > 1)
                    {
                        request.RequestedShipment.MasterTrackingId = new ShipWebServiceClient.ShipServiceWebReference.TrackingId();
                        request.RequestedShipment.MasterTrackingId.TrackingNumber = MasterTrackingNo;
                        request.RequestedShipment.RequestedPackageLineItems[0].SequenceNumber = i.ToString();
                    }

                    ProcessShipmentReply reply = service.processShipment(request);
                    if (reply.HighestSeverity == NotificationSeverityType.SUCCESS ||
                        reply.HighestSeverity == NotificationSeverityType.NOTE || reply.HighestSeverity == NotificationSeverityType.WARNING)
                    {
                        _IsSucceded = true;
                        if (i == 1)// Need to fix here for quantity 2 and Multiline itms issue
                            MasterTrackingNo = reply.CompletedShipmentDetail.MasterTrackingId.TrackingNumber;
                        ShowShipmentReply(isCodShipment, reply, txtLblPath.Text);
                    }
                    else
                    {
                        _IsSucceded = false;
                        _FailReason += reply.Notifications[0].Message;
                    }
                    if (listLineItemSecondValues != null && _IsSucceded)
                    {
                        foreach (SecondaryLineItems LineItemSecondValues in listLineItemSecondValues)
                        {
                            childRequest.RequestedShipment.RequestedPackageLineItems[0].Weight.Value = LineItemSecondValues.Weight;
                            childRequest.RequestedShipment.RequestedPackageLineItems[0].Dimensions.Length = LineItemSecondValues.Length;
                            childRequest.RequestedShipment.RequestedPackageLineItems[0].Dimensions.Width = LineItemSecondValues.Width;
                            childRequest.RequestedShipment.RequestedPackageLineItems[0].Dimensions.Height = LineItemSecondValues.Height;

                            childRequest.RequestedShipment.MasterTrackingId = new ShipWebServiceClient.ShipServiceWebReference.TrackingId();
                            childRequest.RequestedShipment.MasterTrackingId.TrackingNumber = reply.CompletedShipmentDetail.MasterTrackingId.TrackingNumber;
                            childRequest.RequestedShipment.RequestedPackageLineItems[0].SequenceNumber = LineItemSecondValues.SequenceNumber.ToString();
                            ProcessShipmentReply replyChild = service.processShipment(childRequest);
                            if (replyChild.HighestSeverity == NotificationSeverityType.SUCCESS ||
                                replyChild.HighestSeverity == NotificationSeverityType.NOTE || replyChild.HighestSeverity == NotificationSeverityType.WARNING)
                            {
                                _IsSucceded = true;
                                ShowShipmentReply(isCodShipment, replyChild, txtLblPath.Text);
                            }
                            else
                            {
                                _IsSucceded = false;
                                _FailReason += reply.Notifications[0].Message;
                            }
                        }
                    }
                    ShowNotifications(reply);
                }
            }
            catch (SoapException e)
            {
                Console.WriteLine(e.Detail.InnerText);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            //Console.WriteLine("\nPress any key to quit !");
            //Console.ReadKey();
        }

        private WebAuthenticationDetail SetWebAuthenticationDetail()
        {
            WebAuthenticationDetail wad = new WebAuthenticationDetail();
            wad.UserCredential = new WebAuthenticationCredential();
            wad.ParentCredential = new WebAuthenticationCredential();
            wad.UserCredential.Key = "x1pHFuEKcnjBFiKZ";//Q6HmWN7Bsmb3Is9u*/"; // Replace "XXX" with the Key
            wad.UserCredential.Password = "H3CGUKRt8K2FayoICoXU1Ep4Z"; // Replace "XXX" with the Password
            wad.ParentCredential = new WebAuthenticationCredential();

            wad.ParentCredential.Key = "x1pHFuEKcnjBFiKZ";//Q6HmWN7Bsmb3Is9u*/"; // Replace "XXX" with the Key
            wad.ParentCredential.Password = "H3CGUKRt8K2FayoICoXU1Ep4Z"; // Replace "XXX" with the Password

            if (usePropertyFile()) //Set values from a file for testing purposes
            {
                wad.UserCredential.Key = getProperty("key");
                wad.UserCredential.Password = getProperty("password");
                wad.ParentCredential.Key = getProperty("parentkey");
                wad.ParentCredential.Password = getProperty("parentpassword");
            }
            return wad;
        }

        private ProcessShipmentRequest CreateShipmentRequest(bool isCodShipment, DataRow dr)
        {
            // Build the ShipmentRequest
            ProcessShipmentRequest request = new ProcessShipmentRequest();
            //
            request.WebAuthenticationDetail = SetWebAuthenticationDetail();
            //
            request.ClientDetail = new ClientDetail();
            request.ClientDetail.AccountNumber = "642330942"; // Replace "XXX" with the client's account number
            request.ClientDetail.MeterNumber = "111908162"; // Replace "XXX" with the client's meter number
            if (usePropertyFile()) //Set values from a file for testing purposes
            {
                request.ClientDetail.AccountNumber = getProperty("accountnumber");
                request.ClientDetail.MeterNumber = getProperty("meternumber");
            }
            //
            request.TransactionDetail = new TransactionDetail();
            request.TransactionDetail.CustomerTransactionId = "***Ground Domestic Ship Request using VC#***"; // The client will get the same value back in the response
                                                                                                              //
            request.Version = new VersionId();
            //
            SetShipmentDetails(request, dr);
            //
            SetSender(request);
            //
            SetRecipient(request, dr);
            //
            SetPayment(request);
            //
            SetLabelDetails(request);
            //
            SetPackageLineItems(request, dr);
            //
            return request;
        }

        private void SetShipmentDetails(ProcessShipmentRequest request, DataRow dr)
        {
            request.RequestedShipment = new RequestedShipment();
            request.RequestedShipment.ShipTimestamp = DateTime.Now; // Ship date and time
            request.RequestedShipment.ServiceType = ServiceType.FEDEX_GROUND; // Service types are FEDEX_GROUND, GROUND_HOME_DELIVERY ...
            request.RequestedShipment.PackagingType = PackagingType.YOUR_PACKAGING; // Packaging type YOUR_PACKAGING, ...
                                                                                    //
            request.RequestedShipment.PackageCount = dr.ItemArray[20].ToString().Trim();
            // set HAL
            bool isHALShipment = false;
            //if (isHALShipment)
            //    SetHAL(request);
        }

        private void SetSender(ProcessShipmentRequest request)
        {
            request.RequestedShipment.Shipper = new Party();
            request.RequestedShipment.Shipper.Contact = new Contact();
            request.RequestedShipment.Shipper.Contact.PersonName = "US FURNISHING EXPRESS";
            request.RequestedShipment.Shipper.Contact.CompanyName = "US FURNISHINGS EXPRESS CORP";
            request.RequestedShipment.Shipper.Contact.PhoneNumber = "9093641661";
            //
            request.RequestedShipment.Shipper.Address = new Address();
            request.RequestedShipment.Shipper.Address.StreetLines = new string[2] { "5125 SCHAEFER AVE", "STE 104" };
            request.RequestedShipment.Shipper.Address.City = "CHINO";
            request.RequestedShipment.Shipper.Address.StateOrProvinceCode = "CA";
            request.RequestedShipment.Shipper.Address.PostalCode = "91710";
            request.RequestedShipment.Shipper.Address.CountryCode = "US";

        }

        private void SetRecipient(ProcessShipmentRequest request, DataRow dr)
        {
            request.RequestedShipment.Recipient = new Party();
            request.RequestedShipment.Recipient.Contact = new Contact();
            request.RequestedShipment.Recipient.Contact.PersonName = dr.ItemArray[4].ToString().Trim();
            request.RequestedShipment.Recipient.Contact.CompanyName = dr.ItemArray[4].ToString().Trim();
            request.RequestedShipment.Recipient.Contact.PhoneNumber = dr.ItemArray[6].ToString().Trim();
            //
            request.RequestedShipment.Recipient.Address = new Address();
            request.RequestedShipment.Recipient.Address.StreetLines = new string[2] { dr.ItemArray[8].ToString().Trim(), dr.ItemArray[9].ToString().Trim() };
            request.RequestedShipment.Recipient.Address.City = dr.ItemArray[10].ToString().Trim();
            request.RequestedShipment.Recipient.Address.StateOrProvinceCode = dr.ItemArray[11].ToString().Trim();
            request.RequestedShipment.Recipient.Address.PostalCode = dr.ItemArray[12].ToString().Trim();
            request.RequestedShipment.Recipient.Address.CountryCode = "US";
        }

        private void SetPayment(ProcessShipmentRequest request)
        {
            request.RequestedShipment.ShippingChargesPayment = new Payment();
            request.RequestedShipment.ShippingChargesPayment.PaymentType = PaymentType.SENDER;
            request.RequestedShipment.ShippingChargesPayment.Payor = new Payor();
            request.RequestedShipment.ShippingChargesPayment.Payor.ResponsibleParty = new Party();
            request.RequestedShipment.ShippingChargesPayment.Payor.ResponsibleParty.AccountNumber = "642330942";// "801472842"; 
                                                                                                                // Replace "XXX" with client's account number
            if (usePropertyFile()) //Set values from a file for testing purposes
            {
                request.RequestedShipment.ShippingChargesPayment.Payor.ResponsibleParty.AccountNumber = getProperty("payoraccount");
            }
            request.RequestedShipment.ShippingChargesPayment.Payor.ResponsibleParty.Contact = new Contact();
            request.RequestedShipment.ShippingChargesPayment.Payor.ResponsibleParty.Address = new Address();
            request.RequestedShipment.ShippingChargesPayment.Payor.ResponsibleParty.Address.CountryCode = "US";
        }

        private void SetLabelDetails(ProcessShipmentRequest request)
        {
            request.RequestedShipment.LabelSpecification = new LabelSpecification();
            request.RequestedShipment.LabelSpecification.ImageType = ShippingDocumentImageType.PDF; // Image types PDF, PNG, DPL, ...
            request.RequestedShipment.LabelSpecification.ImageTypeSpecified = true;
            request.RequestedShipment.LabelSpecification.LabelFormatType = LabelFormatType.COMMON2D;
        }

        private void SetHAL(ProcessShipmentRequest request)
        {
            request.RequestedShipment.SpecialServicesRequested = new ShipmentSpecialServicesRequested();
            request.RequestedShipment.SpecialServicesRequested.SpecialServiceTypes = new ShipmentSpecialServiceType[1];
            request.RequestedShipment.SpecialServicesRequested.SpecialServiceTypes[0] = ShipmentSpecialServiceType.HOLD_AT_LOCATION;
            //
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail = new HoldAtLocationDetail();
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.PhoneNumber = "9011234567";
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress = new ContactAndAddress();
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Contact = new Contact();
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Contact.PersonName = "Tester";
            //
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Address = new Address();
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Address.StreetLines = new string[1];
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Address.StreetLines[0] = "45 Noblestown Road";
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Address.City = "Pittsburgh";
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Address.StateOrProvinceCode = "PA";
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Address.PostalCode = "15220";
            request.RequestedShipment.SpecialServicesRequested.HoldAtLocationDetail.LocationContactAndAddress.Address.CountryCode = "US";
        }

        private void SetCOD(ProcessShipmentRequest request)
        {
            request.RequestedShipment.RequestedPackageLineItems[0].SpecialServicesRequested = new PackageSpecialServicesRequested();
            request.RequestedShipment.RequestedPackageLineItems[0].SpecialServicesRequested.SpecialServiceTypes = new PackageSpecialServiceType[1];
            request.RequestedShipment.RequestedPackageLineItems[0].SpecialServicesRequested.SpecialServiceTypes[0] = PackageSpecialServiceType.COD;
            //
            request.RequestedShipment.RequestedPackageLineItems[0].SpecialServicesRequested.CodDetail = new CodDetail();
            request.RequestedShipment.RequestedPackageLineItems[0].SpecialServicesRequested.CodDetail.CollectionType = CodCollectionType.GUARANTEED_FUNDS;
            request.RequestedShipment.RequestedPackageLineItems[0].SpecialServicesRequested.CodDetail.CodCollectionAmount = new Money();
            request.RequestedShipment.RequestedPackageLineItems[0].SpecialServicesRequested.CodDetail.CodCollectionAmount.Amount = 250.00M;
            request.RequestedShipment.RequestedPackageLineItems[0].SpecialServicesRequested.CodDetail.CodCollectionAmount.Currency = "USD";
        }

        private void ShowShipmentReply(bool isCodShipment, ProcessShipmentReply reply, string lablePath)
        {
            Console.WriteLine("Shipment Reply details:");
            Console.WriteLine("Package details\n");
            // Details for each package
            foreach (CompletedPackageDetail packageDetail in reply.CompletedShipmentDetail.CompletedPackageDetails)
            {
                ShowTrackingDetails(packageDetail.TrackingIds);
                if (null != packageDetail.PackageRating && null != packageDetail.PackageRating.PackageRateDetails)
                {
                    ShowPackageRateDetails(packageDetail.PackageRating.PackageRateDetails);
                }
                else
                {
                    Console.WriteLine("Rate information not returned");
                }
                ShowBarcodeDetails(packageDetail.OperationalDetail.Barcodes);
                ShowShipmentLabels(isCodShipment, reply.CompletedShipmentDetail, packageDetail, lablePath);
            }
            ShowPackageRouteDetails(reply.CompletedShipmentDetail.OperationalDetail);
        }

        private void ShowShipmentLabels(bool isCodShipment, CompletedShipmentDetail completedShipmentDetail, CompletedPackageDetail packageDetail
            , string lablePath)
        {
            if (null != packageDetail.Label.Parts[0].Image)
            {
                // Save outbound shipping label
                string LabelPath = lablePath.Trim();
                if (usePropertyFile())
                {
                    LabelPath = getProperty("labelpath");
                }
                string LabelFileName = LabelPath + packageDetail.TrackingIds[0].TrackingNumber + ".pdf";
                SaveLabel(LabelFileName, packageDetail.Label.Parts[0].Image);
            }
        }

        private void ShowTrackingDetails(TrackingId[] TrackingIds)
        {
            // Tracking information for each package
            Console.WriteLine("Tracking details");
            if (TrackingIds != null)
            {
                for (int i = 0; i < TrackingIds.Length; i++)
                {
                    TrackingId += TrackingIds[i].TrackingNumber.ToString() + ", ";
                    Console.WriteLine("Tracking # {0} Form ID {1}", TrackingIds[i].TrackingNumber, TrackingIds[i].FormId);
                }
            }
        }

        private void ShowPackageRateDetails(PackageRateDetail[] PackageRateDetails)
        {
            foreach (PackageRateDetail ratedPackage in PackageRateDetails)
            {
                Console.WriteLine("\nRate details");
                if (ratedPackage.BillingWeight != null)
                    Console.WriteLine("Billing weight {0} {1}", ratedPackage.BillingWeight.Value, ratedPackage.BillingWeight.Units);
                if (ratedPackage.BaseCharge != null)
                    Console.WriteLine("Base charge {0} {1}", ratedPackage.BaseCharge.Amount, ratedPackage.BaseCharge.Currency);
                if (ratedPackage.TotalSurcharges != null)
                    Console.WriteLine("Total surcharge {0} {1}", ratedPackage.TotalSurcharges.Amount, ratedPackage.TotalSurcharges.Currency);
                if (ratedPackage.Surcharges != null)
                {
                    // Individual surcharge for each package
                    foreach (Surcharge surcharge in ratedPackage.Surcharges)
                        Console.WriteLine(" {0} surcharge {1} {2}", surcharge.SurchargeType, surcharge.Amount.Amount, surcharge.Amount.Currency);
                }
                if (ratedPackage.NetCharge != null)
                    Console.WriteLine("Net charge {0} {1}", ratedPackage.NetCharge.Amount, ratedPackage.NetCharge.Currency);
            }
        }

        private void ShowBarcodeDetails(PackageBarcodes barcodes)
        {
            // Barcode information for each package
            Console.WriteLine("\nBarcode details");
            if (barcodes != null)
            {
                if (barcodes.StringBarcodes != null)
                {
                    for (int i = 0; i < barcodes.StringBarcodes.Length; i++)
                    {
                        Console.WriteLine("String barcode {0} Type {1}", barcodes.StringBarcodes[i].Value, barcodes.StringBarcodes[i].Type);
                    }
                }

                if (barcodes.BinaryBarcodes != null)
                {
                    for (int i = 0; i < barcodes.BinaryBarcodes.Length; i++)
                    {
                        Console.WriteLine("Binary barcode Type {0}", barcodes.BinaryBarcodes[i].Type);
                    }
                }
            }
        }

        private void ShowPackageRouteDetails(ShipmentOperationalDetail routingDetail)
        {
            Console.WriteLine("\nRouting details");
            Console.WriteLine("URSA prefix {0} suffix {1}", routingDetail.UrsaPrefixCode, routingDetail.UrsaSuffixCode);
            Console.WriteLine("Service commitment {0} Airport ID {1}", routingDetail.DestinationLocationId, routingDetail.AirportId);

            if (routingDetail.DeliveryDaySpecified)
            {
                Console.WriteLine("Delivery day " + routingDetail.DeliveryDay);
            }
            if (routingDetail.DeliveryDateSpecified)
            {
                Console.WriteLine("Delivery date " + routingDetail.DeliveryDate.ToShortDateString());
            }
            if (routingDetail.TransitTimeSpecified)
            {
                Console.WriteLine("Transit time " + routingDetail.TransitTime);
            }
        }

        private void SaveLabel(string labelFileName, byte[] labelBuffer)
        {
            // Save label buffer to file
            FileStream LabelFile = new FileStream(labelFileName, FileMode.Create);
            LabelFile.Write(labelBuffer, 0, labelBuffer.Length);
            LabelFile.Close();
            // Display label in Acrobat
            //  DisplayLabel(labelFileName);
        }

        private void DisplayLabel(string labelFileName)
        {
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo(labelFileName);
            info.UseShellExecute = true;
            info.Verb = "open";
            System.Diagnostics.Process.Start(info);
        }

        private void ShowNotifications(ProcessShipmentReply reply)
        {
            Console.WriteLine("Notifications");
            for (int i = 0; i < reply.Notifications.Length; i++)
            {
                Notification notification = reply.Notifications[i];
                Console.WriteLine("Notification no. {0}", i);
                Console.WriteLine(" Severity: {0}", notification.Severity);
                Console.WriteLine(" Code: {0}", notification.Code);
                Console.WriteLine(" Message: {0}", notification.Message);
                Console.WriteLine(" Source: {0}", notification.Source);
            }
        }

        private bool usePropertyFile() //Set to true for common properties to be set with getProperty function.
        {
            return getProperty("usefile").Equals("True");
        }

        private String getProperty(String propertyname) //Sets common properties for testing purposes.
        {
            try
            {
                String filename = "C:\\filepath\\filename.txt";
                if (System.IO.File.Exists(filename))
                {
                    System.IO.StreamReader sr = new System.IO.StreamReader(filename);
                    do
                    {
                        String[] parts = sr.ReadLine().Split(',');
                        if (parts[0].Equals(propertyname) && parts.Length == 2)
                        {
                            return parts[1];
                        }
                    }
                    while (!sr.EndOfStream);
                }
                Console.WriteLine("Property {0} set to default 'XXX'", propertyname);
                return "XXX";
            }
            catch (Exception e)
            {
                Console.WriteLine("Property {0} set to default 'XXX'", propertyname);
                return "XXX";
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        public bool SetPackageLineItems(ProcessShipmentRequest request, DataRow dr)
        {
            bool _IsLineItemFound = false;

            foreach (DataRow Row in dtLineItemData.Rows)
            {
                if (Row.ItemArray[0].ToString().Trim() == dr.ItemArray[17].ToString().Trim())
                {
                    if (Convert.ToInt32(Row.ItemArray[1]) > 1)
                    {
                        childRequest = request;
                        request.RequestedShipment.PackageCount = Row.ItemArray[1].ToString();
                    }
                    _IsLineItemFound = true;
                    request.RequestedShipment.RequestedPackageLineItems = new RequestedPackageLineItem[1];
                    request.RequestedShipment.RequestedPackageLineItems[0] = new RequestedPackageLineItem();
                    request.RequestedShipment.RequestedPackageLineItems[0].SequenceNumber = "1";
                    // Package weight information
                    request.RequestedShipment.RequestedPackageLineItems[0].Weight = new Weight();
                    request.RequestedShipment.RequestedPackageLineItems[0].Weight.Value = Row.ItemArray[5].GetType().Name != "DBNull" ? Convert.ToInt32(Row.ItemArray[5]) : 0;
                    request.RequestedShipment.RequestedPackageLineItems[0].Weight.Units = WeightUnits.LB;
                    //
                    request.RequestedShipment.RequestedPackageLineItems[0].Dimensions = new Dimensions();
                    request.RequestedShipment.RequestedPackageLineItems[0].Dimensions.Length = Row.ItemArray[2]?.ToString();
                    request.RequestedShipment.RequestedPackageLineItems[0].Dimensions.Width = Row.ItemArray[3]?.ToString();
                    request.RequestedShipment.RequestedPackageLineItems[0].Dimensions.Height = Row.ItemArray[4]?.ToString();
                    request.RequestedShipment.RequestedPackageLineItems[0].Dimensions.Units = LinearUnits.IN;
                    // Reference details
                    request.RequestedShipment.RequestedPackageLineItems[0].CustomerReferences = new CustomerReference[3] { new CustomerReference(), new CustomerReference(), new CustomerReference() };
                    request.RequestedShipment.RequestedPackageLineItems[0].CustomerReferences[0].CustomerReferenceType = CustomerReferenceType.CUSTOMER_REFERENCE;
                    request.RequestedShipment.RequestedPackageLineItems[0].CustomerReferences[0].Value = dr.ItemArray[0]?.ToString().Trim();
                    request.RequestedShipment.RequestedPackageLineItems[0].CustomerReferences[1].CustomerReferenceType = CustomerReferenceType.INVOICE_NUMBER;
                    request.RequestedShipment.RequestedPackageLineItems[0].CustomerReferences[1].Value = "";
                    request.RequestedShipment.RequestedPackageLineItems[0].CustomerReferences[2].CustomerReferenceType = CustomerReferenceType.P_O_NUMBER;
                    request.RequestedShipment.RequestedPackageLineItems[0].CustomerReferences[2].Value = dr.ItemArray[0]?.ToString().Trim();

                    if (Convert.ToInt32(Row.ItemArray[1]) > 1)
                    {
                        listLineItemSecondValues = new List<SecondaryLineItems>();
                        for (int i = 0; i < Convert.ToInt32(Row.ItemArray[1]) - 1; i++)
                        {
                            listLineItemSecondValues.Add(
                         new SecondaryLineItems
                         {
                             Weight = Row.ItemArray[9 + (i * 4)].GetType().Name != "DBNull" ? Convert.ToInt32(Row.ItemArray[9 + (i * 4)]) : 0,
                             Length = Row.ItemArray[6 + (i * 4)]?.ToString(),
                             Width = Row.ItemArray[7 + (i * 4)]?.ToString(),
                             Height = Row.ItemArray[8 + (i * 4)]?.ToString(),
                             SequenceNumber = i + 2
                         });
                        }
                    }
                    break;
                }
            }
            if (!_IsLineItemFound)
                _FailReason = "Ohoo.... Line Item Not Found!!";
            return _IsLineItemFound;
        }
    }
}
