using Newtonsoft.Json;
using OfficeOpenXml;
using sapfewse;
using saprotwr.net;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace asmu
{
    internal class Program
    {
        public static string FAR = @"D:\Aghils\ASSETS\FAR.xlsx";
        public static string EQM = @"D:\Aghils\ASSETS\EQM.xlsx";
        public static string REPORT = @"D:\Aghils\ASSETS\ASSET MASTER UPDATE.xlsx";
        public static IDictionary<string, string> SubTypeCollections;

        public static List<LoggingSystem> log;

        //      public static String REPORT = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ASSET MASTER UPDATE.xlsx");
        public static List<Equipment> EquipmentMaster = new List<Equipment>();
        public static List<Equipment> AssetMaster = new List<Equipment>();
        public static List<Equipment> EquipmentMasterOnServer = new List<Equipment>();
        public static List<Equipment> Equipments_tobe_updated_to_Server = new List<Equipment>();

        private static GuiSession session;


        private static void Main(string[] args)
        {
            Console.WriteLine("------------------------------------------------------------------------");
            Console.WriteLine("----------------EMMA CONSOLE @ AGHIL K MOHAN 2018-----------------------");
            Console.WriteLine("------------------------------------------------------------------------");
            Console.WriteLine();
            log = new List<LoggingSystem>();
            cmd();
        }

        private static void cmd()
        {
            Console.WriteLine();
            var command = Console.ReadLine().ToLower();
            switch (command)
            {
                default:
                    Console.WriteLine("PLEASE ENTER A VALID COMMAND");

                    break;

                case "download":
                    CheckForChanges();
                    UpdateAssetMaster();
                    break;

                case "sync":
                    ReadASM();
                    GetDataFromServer();

                    break;

                case "upload":
                    ReadASM();
                    CheckForChanges();
                    UpdateChanges();
                    break;

                case "cls":
                    Console.Clear();
                    break;
                case "wipe":
                    CheckForChanges();
                    WipeASM();
                    break;

                case "open":
                    OpenASM();
                    break;

                case "close":
                    Environment.Exit(0);
                    break;

                case "read":
                    ReadASM();
                    break;

                case "write":
                    WriteToFile();
                    break;

                case "check":
                    CheckForChanges();
                    break;

                case "fb":

                    break;

            }

            cmd();
        }

        private static void CrossCheckDataFromServer()
        {
            Console.WriteLine("Starting Cross Check for " + AssetMaster.Count);
            int changes = 0;
            for (int i = 0; i < AssetMaster.Count; i++)
            {
                try
                {
                    Console.Write("\rChecking " + i + " of " + AssetMaster.Count + " items. " + changes + " Updates Found");
                    Equipment equipment = AssetMaster[i];
                    Equipment e_onServer = EquipmentMasterOnServer.Where((e) => e.EquipmentNumber == equipment.EquipmentNumber).First();
                    if (e_onServer != null)
                    {
                        if (e_onServer.New.EquipmentDescription != equipment.New.EquipmentDescription)
                        {
                            e_onServer.New.EquipmentDescription = equipment.New.EquipmentDescription;
                            Equipments_tobe_updated_to_Server.Add(e_onServer);
                            //Console.WriteLine(e_onServer.EquipmentDescription + "-----" + equipment.New.EquipmentDescription);
                            changes++;
                        }
                    }
                }
                catch (Exception e)
                {

                    // Console.WriteLine(e.Message);
                }

            }

            Console.WriteLine("\nCross Check Completed.");
            UpdateChangesToServer();
        }

        private static void UpdateChangesToServer()
        {
            Console.WriteLine("\nUpdating Changes to Server");
            foreach (Equipment equipment in Equipments_tobe_updated_to_Server)
            {
                using (HttpClient c = new HttpClient())
                {
                    string jso = JsonConvert.SerializeObject(equipment).ToString();
                    var content = new StringContent(jso, Encoding.UTF8, "application/json");
                    Console.Write(jso);
                    Console.ReadLine();
                    var resp = c.PutAsync(@"http://xo.rs/api/Equipments" + equipment.EquipmentNumber, content);
                    Console.WriteLine(resp.Result.StatusCode);
                }
            }
        }

        private static async Task GetDataFromServer()
        {
            Console.WriteLine();
            Console.WriteLine("Downloading Data......");
            using (HttpClient client = new HttpClient())
            {
                string response = await client.GetStringAsync(@"http://xo.rs/api/Equipments/");
                EquipmentMasterOnServer = JsonConvert.DeserializeObject<List<Equipment>>(response);

            }

            Console.WriteLine(EquipmentMasterOnServer.Count + " Equipments Found on Server");

            CrossCheckDataFromServer();
        }

        private static bool CheckForChanges()
        {
            var changes = 0;
            Console.WriteLine("Checking Asset Master for Changes...");
            for (var i = 0; i < AssetMaster.Count; i++)
            {
                if (AssetMaster[i].Old.AssetDescription != AssetMaster[i].New.AssetDescription)
                    changes++;

                if (AssetMaster[i].Old.EquipmentDescription != AssetMaster[i].New.EquipmentDescription)
                    changes++;

                if (AssetMaster[i].Old.OperationId != AssetMaster[i].New.OperationId)
                    changes++;

                if (AssetMaster[i].Old.ModelNumber != AssetMaster[i].New.ModelNumber)
                    changes++;

                if (AssetMaster[i].Old.SerialNumber != AssetMaster[i].New.SerialNumber)
                    changes++;

                if (AssetMaster[i].Old.Dimensions != AssetMaster[i].New.Dimensions)
                    changes++;

                //if (AssetMaster[i].Old.Weight != AssetMaster[i].New.Weight)
                //    changes++;

                if (AssetMaster[i].Old.SubType != AssetMaster[i].New.SubType)
                    changes++;

                Console.Write("\r{0} New Changes Found   ", changes);
            }

            if (changes > 0)
                return true;
            return false;
        }

        private static void ReadASM()
        {
            AssetMaster = new List<Equipment>();
            var fi = new FileInfo(REPORT);
            // FileStream stream = File.Open(REPORT,FileMode.Open);
            using (var excelPackage = new ExcelPackage(fi))
            {
                var myWorkbook = excelPackage.Workbook;
                var myWorksheet = myWorkbook.Worksheets[1];
                Console.WriteLine();
                for (var i = 2; i < myWorksheet.Dimension.End.Row; i++)
                {
                    var e = new Equipment();
                    e.AssetNumber = myWorksheet.Cells[i, 1].Text.Trim();
                    e.EquipmentNumber = myWorksheet.Cells[i, 2].Text.Trim();
                    e.Old.AssetDescription = myWorksheet.Cells[i, 3].Text.Trim();
                    e.New.AssetDescription = myWorksheet.Cells[i, 4].Text.Trim();
                    e.Old.EquipmentDescription = myWorksheet.Cells[i, 5].Text.Trim();
                    e.New.EquipmentDescription = myWorksheet.Cells[i, 6].Text.Trim();
                    e.Old.OperationId = myWorksheet.Cells[i, 7].Text.Trim();
                    e.New.OperationId = myWorksheet.Cells[i, 8].Text.Trim();
                    e.Old.SubTypeDescription = myWorksheet.Cells[i, 9].Text.Trim();
                    e.New.SubTypeDescription = myWorksheet.Cells[i, 11].Text.Trim();
                    e.Old.SubType = myWorksheet.Cells[i, 10].Text.Trim();
                    e.Old.Weight = myWorksheet.Cells[i, 12].Text.Trim();
                    e.Old.WeightUnit = myWorksheet.Cells[i, 13].Text.Trim();
                    e.New.Weight = myWorksheet.Cells[i, 14].Text.Trim();
                    e.New.WeightUnit = myWorksheet.Cells[i, 15].Text.Trim();
                    e.Old.Dimensions = myWorksheet.Cells[i, 16].Text.Trim();
                    e.New.Dimensions = myWorksheet.Cells[i, 17].Text.Trim();
                    e.Old.ModelNumber = myWorksheet.Cells[i, 22].Text.Trim();
                    e.New.ModelNumber = myWorksheet.Cells[i, 23].Text.Trim();
                    e.Old.SerialNumber = myWorksheet.Cells[i, 24].Text.Trim();
                    e.New.SerialNumber = myWorksheet.Cells[i, 25].Text.Trim();
                    e.BookValue = myWorksheet.Cells[i, 31].Text.Trim();
                    e.AcquisitionValue = myWorksheet.Cells[i, 30].Text.Trim();
                    e.AcquisitionDate = myWorksheet.Cells[i, 29].Formula.Trim();
                    e.Old.AssetLocation = myWorksheet.Cells[i, 26].Text.Trim();
                    e.Old.EquipmentLocation = myWorksheet.Cells[i, 27].Text.Trim();
                    e.New.EquipmentLocation = myWorksheet.Cells[i, 28].Text.Trim();


                    Console.Write("\r{0} Equipments Found   ", AssetMaster.Count);

                    AssetMaster.Add(e);
                }

                Console.Write(" Process Completed   ");
            }
        }

        private static void UpdateChanges()
        {
            UpdateSubtypeCollections();
            ActivateSAP();
            Console.WriteLine();
            for (var i = 0; i < AssetMaster.Count; i++)
            {
                try
                {


                    if (AssetMaster[i].Old.EquipmentDescription != AssetMaster[i].New.EquipmentDescription)
                        eUpdateEquipmentDescription(i);

                    if (AssetMaster[i].Old.OperationId != AssetMaster[i].New.OperationId) EUpdateOperationId(i);
                    if (AssetMaster[i].Old.ModelNumber != AssetMaster[i].New.ModelNumber) EUpdateModelNumber(i);
                    if (AssetMaster[i].Old.SerialNumber != AssetMaster[i].New.SerialNumber) EUpdateSerialNumber(i);
                    if (AssetMaster[i].Old.SubTypeDescription != AssetMaster[i].New.SubTypeDescription) EUpdateSubType(i);
                    if (AssetMaster[i].Old.EquipmentLocation != AssetMaster[i].New.EquipmentLocation && AssetMaster[i].New.EquipmentLocation != string.Empty) EUpdateEquipmentLocation(i);


                    // if (AssetMaster[i].Old.AssetDescription.Trim() != AssetMaster[i].New.EquipmentDescription.Trim() && AssetMaster[i].New.EquipmentDescription != "") EUpdateAssetDescription(i);
                    Console.Write("\r{0} Records Updated in SAP  ", i);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);

                }
            }

            WriteToFile();
            updateLog();
        }

        private static void EUpdateEquipmentLocation(int i)
        {

            ((GuiMainWindow)session.FindById("wnd[0]")).Maximize();
            ((GuiOkCodeField)session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/niw21";
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById("wnd[0]/usr/ctxtRIWO00-QMART")).Text = "ZS";
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7710/txtRIWO00-HEADKTXT")).Text = "EQUIPMENT MOVEMENT";
            ((GuiTextField)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7330/ctxtVIQMEL-STRMN")).Text = DateTime.Today.ToString("dd.MM.yyyy");
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiComboBox)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7900/ssubUSER0001:SAPLXQQM:0105/cmbZCBS_ETM_EMR_H-RECEIVER_TYP")).Key = GetLocationType(AssetMaster[i].New.EquipmentLocation);
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiCTextField)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7900/ssubUSER0001:SAPLXQQM:0105/ctxtZCBS_ETM_EMR_H-REFERENCE_R")).Text = AssetMaster[i].New.EquipmentLocation;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiComboBox)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7900/ssubUSER0001:SAPLXQQM:0105/cmbZCBS_ETM_EMR_H-SENDER_TYPE")).Key = GetLocationType(AssetMaster[i].Old.EquipmentLocation);
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiCTextField)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7900/ssubUSER0001:SAPLXQQM:0105/ctxtZCBS_ETM_EMR_H-REFERENCE_S")).Text = AssetMaster[i].Old.EquipmentLocation;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);

            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTab)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB19")).Select();
            ((GuiGridView)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB19/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/ssubUSER0001:SAPLXQQM:0103/cntlALV_GRID_CONT/shellcont/shell")).ModifyCell(0, "EQUNR", AssetMaster[i].EquipmentNumber);
            ((GuiGridView)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB19/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/ssubUSER0001:SAPLXQQM:0103/cntlALV_GRID_CONT/shellcont/shell")).TriggerModified();
            ((GuiGridView)session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB19/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/ssubUSER0001:SAPLXQQM:0103/cntlALV_GRID_CONT/shellcont/shell")).PressEnter();
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(11);

            if (!((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text.Contains("saved"))
            {
                var l = new LoggingSystem();
                l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
                l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
                l.Type = "EMR FAIL";
                l.OldValue = AssetMaster[i].Old.EquipmentLocation;
                l.NewValue = AssetMaster[i].New.EquipmentLocation;
                //AssetMaster[i].Old.EquipmentLocation = AssetMaster[i].New.EquipmentLocation;

                log.Add(l);
                // Console.ReadLine();
            }
            else
            {
                ((GuiOkCodeField)session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/niw22";
                ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
                ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
                ((GuiButton)session.FindById("wnd[0]/tbar[1]/btn[13]")).Press();
                ((GuiButton)session.FindById("wnd[0]/tbar[1]/btn[16]")).Press();
                if (((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text.Contains("Reference date for completion will be determined by notification type"))
                    ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);


                if (((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text.Contains("completed"))
                {
                    var l = new LoggingSystem();
                    l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
                    l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
                    l.Type = "EMR SUCCESS";
                    l.OldValue = AssetMaster[i].Old.EquipmentLocation;
                    l.NewValue = AssetMaster[i].New.EquipmentLocation;
                    AssetMaster[i].Old.EquipmentLocation = AssetMaster[i].New.EquipmentLocation;

                    log.Add(l);
                }
                else
                {
                    var l = new LoggingSystem();
                    l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
                    l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
                    l.Type = "EMR FAIL";
                    l.OldValue = AssetMaster[i].Old.EquipmentLocation;
                    l.NewValue = AssetMaster[i].New.EquipmentLocation;
                    //AssetMaster[i].Old.EquipmentLocation = AssetMaster[i].New.EquipmentLocation;

                    log.Add(l);
                }


            }




        }

        private static string GetLocationType(string equipmentLocation)
        {
            if (equipmentLocation.Contains("BS-"))
            {
                return "B";
            }
            else
            {
                return "P";
            }
        }

        private static void UpdateSubtypeCollections()
        {
            SubTypeCollections = new Dictionary<string, string>()
        {
        {"11AAA","LAND"},
{"11BAA","BUILDING"},
{"11CAA","WORKSHOP"},
{"12AAA","PDS"},
{"12AAB","SPS"},
{"12AAC","SMAPLER CLOCK"},
{"12AAD","SWIVEL JOINTS CONNECTOR"},
{"12AAE","SPS MK2"},
{"12ABA","PGR"},
{"12ABB","QPC"},
{"12ABC","CFBM"},
{"12ABD","MPL"},
{"12ABE","UMT"},
{"12ABF","MBH"},
{"12ABG","PRC"},
{"12ABH","PSC"},
{"12ABI","ILS"},
{"12ABJ","FDI"},
{"12ABK","PKJ"},
{"12ABL","PSJ"},
{"12ABM","FDR"},
{"12ABN","PDC"},
{"12ABO","CTF"},
{"12ABP","CFSM"},
{"12ABQ","CFJ"},
{"12ABR","NTO"},
{"12ABS","BUL"},
{"12ABT","ASB"},
{"12ABU","AGS"},
{"12ABV","RADIO ACTIVE"},
{"12ACA","MIT"},
{"12ADA","SIT PRODUCTION VALVE"},
{"12ADB","SIT DRIVER SECTION"},
{"12ADC","EQUALIZING CROSS OVER"},
{"12AEA","QUARTZ"},
{"12AEB","SAPPHIRE"},
{"12AFA","HEAT GAUGES"},
{"12AGA","DOWN HOLE STRING & ACCESSORIES"},
{"12AHA","PRESSURE ELEMENTS"},
{"12AHB","RECORDING ELEMENTS"},
{"12AHC","RT-7"},
{"12AHD","MECHANICAL CLOCKS"},
{"12AIA","ECHOMETER WELL ANALYZER"},
{"12AIC","ECHOMETER"},
{"12AIB","ECHOMETER GAS GUN"},
{"12BAA","DTR"},
{"12BAB","MIP"},
{"12BAC","UMU"},
{"12BAD","KIT SERVICE TOOLS"},
{"12BAE","GAMMA CAL JIG"},
{"12BBA","BOTTLES HEATING JACKETS"},
{"12BBB","SAMPLER HEATING JACKET"},
{"12BBC","MFTB"},
{"12BBD"," PDS BOTTLES"},
{"12BBE","SPS BOTTLES"},
{"12BBF","NITRUGEN BOOSTER"},
{"12BBG","PDS/SPS OPERATING TOOLS"},
{"12BBH","BHS HAND TOOL BOX"},
{"12BCA","CALIBRATION UNITS"},
{"12BCB","PANELS"},
{"12BDA","MULRI FINGER IMAGING KIT AND REDRESS CALIBRATION JIG"},
{"12BEA","SOFTWARES AND DONGLES"},
{"13AAA","ROPE SOCKETS "},
{"13AAB","STEM BARS"},
{"13AAC","KNUCKLE JOINTS"},
{"13AAD","SWIVEL JOINTS"},
{"13AAE","HIGH DEVEIATION TOOLS "},
{"13AAF","ACCELERATORS "},
{"13AAG","SHOCK ABSORVERS"},
{"13AAH","CENTRALIZERS "},
{"13AAI","STRING X.OVER"},
{"13ABA","MECHANICAL JARS"},
{"13ABB","HYD JARS"},
{"13ABC","KNUCKLE JARS"},
{"13ABD","TUBULAR JARS "},
{"13ABE","SPRING JARS"},
{"13ACA","PLUG SUBS AND PRONGS"},
{"13ACB","TESTING VALVES"},
{"13ACC","SAFETY VALVES"},
{"13ACD","SEPERATION SLEVES"},
{"13ACE","LOCKS"},
{"13ADA","HANGERS"},
{"13AEA","GAS LIFT VALVES"},
{"13AEB","KICK OVER TOOLS"},
{"13AEC","SPACERS"},
{"13AFA","PULLING & RUNNING TOOLS"},
{"13AGA","GO DEVIL"},
{"13AGB","WIRE CUTTER"},
{"13AGC","OVERSHOT"},
{"13AGD","SPEARS"},
{"13AGE","BLIND BOX"},
{"13AGF","IMPRESSION BLOCK"},
{"13AGG","WIRE FINDERS"},
{"13AGH","GRABS"},
{"13AGI","MAGNETS "},
{"13AHA","SELECTIVE SHIFTING TOOLS"},
{"13AHB","NON SELECTIVE SHIFTING TOOLS"},
{"13AIA","GAUGE CUTTERS"},
{"13AIB","TUBING BROACHES"},
{"13AIC","SCRATCHERS"},
{"13AID","SWAGGING TOOLS"},
{"13AIE","BAILERS AND DUMBERS"},
{"13AIF","TUBING END LOCATORS"},
{"13AIG","TUUBING PERFORATORS "},
{"13AIH","ANTI BLOW OUT TOOLS"},
{"13BAA","OFFSHORE UNITS"},
{"13BAB","ONSHORE UNITS"},
{"13BAC","PUMPING UNITS"},
{"13BAD","SPOOLING UNITS"},
{"13BCA","HYD GINPOLE"},
{"13BCB","MANUAL GINPOLE"},
{"13BDA","PRESSURE TEST PANELS"},
{"13BDB","BOP AND VALVES CONTROL PANELS"},
{"13BDC","GRASE AND HYD CONTROL PANELS"},
{"13BEA","POWERPACK"},
{"13BFA","TOOL BOX"},
{"13BGA","BASKETS"},
{"13BHA","CRANES"},
{"13BIA","SLINGS "},
{"13BIB"," MANUAL HYDRAULIC FORK LIFTS"},
{"13BIC","TROLLEYS"},
{"13BJA","WIRELINE ACCESSORIES"},
{"13CAA","STUFFING BOX"},
{"13CBA","LUBRICATOR"},
{"13CCA","GREASE INJECTION HEAD "},
{"13CDA","INJECTION SUB"},
{"13CEA","QUICK TEST SUB"},
{"13CFA","LINE WIPER"},
{"13CGA","TOOL CATCHER"},
{"13CHA","BOP"},
{"13CIA","WELL HEAD ADAPTERS"},
{"13CJA","TOOL TRAP"},
{"13CKA","INHIBITOR SUB"},
{"13CLA","PUMP IN SUB "},
{"14AAA","COMPUTERS"},
{"14ABA","PROJECTORS "},
{"14ACA","TELEVESIONS"},
{"14ADA","MOBILE PHONES"},
{"14AEA","LOAD CELLS "},
{"14AEB","CRANE PANELS"},
{"14BAA","OFFICE - CAMP FURNITURES"},
{"14BBA","WORKSHOP RELATED FURNITURES"},
{"14BCA","KITCHEN FIXTURES"},
{"16AAA","SEDAN CARS"},
{"16BAA","4X4 CARS"},
{"16CAA","PICKUPS"},
{"16DAA","TRUCK UNIT"},
{"16EAA","TRUCK CRANE"},
{"16FAA","FLAT BED TRUCK "},
{"16GAA","TANKER TRUCK"},
{"16HAA","FORKLIFT"},
{"17AAA","BREATHING APPARATUS"},
{"17BAA","GAS DETECTORS"},
{"17CAA","IVMS-VMD"},
{"17DAA","SCAFFOLDING"},
{"18AAA","DIESEL FUEL TANK"},
{"18ABA","PERTOL FUEL TANK"},
{"18BAA","WATER TANKS"},
{"18CAA","CHIMICAL TANKS "},
{"110AA","GENERATORS"},
{"110BA","COMPRESSORS"},
{"110CAA","FUEL PUMP"},
{"110CBA","WATER PUMP"},
{"110DA","FUEL STATIONS"},
{"110EA","FANS"},
{"110FA","OTHERS"},


        };
        }

        private static void EUpdateSubType(int i)
        {
            if (SubTypeCollections.FirstOrDefault(x => x.Value == AssetMaster[i].New.SubTypeDescription).Key == string.Empty)
            {
                Console.WriteLine("------------------------invalid type-----------" + AssetMaster[i].New.SubTypeDescription);
            }
            else
            {
                ((GuiMainWindow)session.FindById("wnd[0]")).Maximize();
                ((GuiOkCodeField)session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nie02";
                ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
                ((GuiTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).Text = AssetMaster[i].EquipmentNumber;
                ((GuiCTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).CaretPosition = 8;
                ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);


                ((GuiTextField)(session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102B:SAPLITO0:1080/ssubXUSR1080:SAPLXTOB:1000/ctxtEQUI-ZZITO_SUB_TYPE"))).Text = SubTypeCollections.FirstOrDefault(x => x.Value == AssetMaster[i].New.SubTypeDescription).Key;
                ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
                ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(11);


                var l = new LoggingSystem();
                l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
                l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
                l.Type = "SUB TYPE";
                l.OldValue = AssetMaster[i].Old.SubTypeDescription;
                l.NewValue = AssetMaster[i].New.SubTypeDescription;
                AssetMaster[i].Old.SubTypeDescription = AssetMaster[i].New.SubTypeDescription;

                log.Add(l);
            }



        }


        private static void EUpdateAssetDescription(int i)
        {
            AssetMaster[i].New.AssetDescription = AssetMaster[i].New.EquipmentDescription;

            ((GuiMainWindow)session.FindById("wnd[0]")).Maximize();
            ((GuiOkCodeField)session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/NAS02";
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById("wnd[0]/usr/ctxtANLA-ANLN1")).Text = AssetMaster[i].AssetNumber;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById(
                    "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-TXT50")
                ).Text = AssetMaster[i].New.AssetDescription; ;
            ((GuiTextField)session.FindById(
                    "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-TXA50")
                ).Text = AssetMaster[i].New.AssetDescription;
            ((GuiTextField)session.FindById(
                    "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLH-ANLHTXT")
                ).Text = AssetMaster[i].New.AssetDescription;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(11);


            var l = new LoggingSystem();
            l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
            l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
            l.Type = "ASSET DESCRIPTION";
            l.OldValue = AssetMaster[i].Old.AssetDescription;
            l.NewValue = AssetMaster[i].New.AssetDescription;
            AssetMaster[i].Old.AssetDescription = AssetMaster[i].New.AssetDescription;

            log.Add(l);
        }

        private static void updateLog()
        {
            var fi = new FileInfo(REPORT);
            using (var excelPackage = new ExcelPackage(fi))
            {
                var myWorkbook = excelPackage.Workbook;
                var myWorksheet = myWorkbook.Worksheets[2];
                var p = myWorksheet.Dimension.End.Row + 1;
                for (var i = 0; i < log.Count; i++)
                {
                    myWorksheet.Cells[p, 1].Value = DateTime.Now;
                    myWorksheet.Cells[p, 2].Value = log[i].EquipmentNumber;
                    myWorksheet.Cells[p, 3].Value = log[i].Type;
                    myWorksheet.Cells[p, 4].Value = log[i].Message;
                    myWorksheet.Cells[p, 5].Value = log[i].OldValue;
                    myWorksheet.Cells[p, 6].Value = log[i].NewValue;
                    Console.Write("\r{0} Logs Updated  ", i);
                    p++;
                }

                excelPackage.Save();
            }
        }

        private static void EUpdateSerialNumber(int i)
        {
            ((GuiMainWindow)session.FindById("wnd[0]")).Maximize();
            ((GuiOkCodeField)session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nie02";
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).Text = AssetMaster[i].EquipmentNumber;
            ((GuiCTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).CaretPosition = 8;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById(
                    "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102D:SAPLITO0:1022/txtITOB-SERGE")
                )
                .Text = AssetMaster[i].New.SerialNumber;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(11);
            AssetMaster[i].Old.SerialNumber = AssetMaster[i].New.SerialNumber;
            var l = new LoggingSystem();
            l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
            l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
            l.Type = "SERIAL NUMBER";
            l.OldValue = AssetMaster[i].Old.SerialNumber;
            l.NewValue = AssetMaster[i].New.SerialNumber;

            log.Add(l);
        }

        private static void EUpdateModelNumber(int i)
        {
            ((GuiMainWindow)session.FindById("wnd[0]")).Maximize();
            ((GuiOkCodeField)session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nie02";
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).Text = AssetMaster[i].EquipmentNumber;
            ((GuiCTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).CaretPosition = 8;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById(
                    "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102D:SAPLITO0:1022/txtITOB-TYPBZ")
                )
                .Text = AssetMaster[i].New.ModelNumber;

            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(11);


            var l = new LoggingSystem();
            l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
            l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
            l.Type = "MODEL NUMBER";
            l.OldValue = AssetMaster[i].Old.ModelNumber;
            l.NewValue = AssetMaster[i].New.ModelNumber;
            AssetMaster[i].Old.ModelNumber = AssetMaster[i].New.ModelNumber;
            log.Add(l);
        }

        private static void EUpdateOperationId(int i)
        {
            ((GuiMainWindow)session.FindById("wnd[0]")).Maximize();
            ((GuiOkCodeField)session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nie02";
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).Text = AssetMaster[i].EquipmentNumber;
            ((GuiCTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).CaretPosition = 8;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById(
                    "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102A:SAPLITO0:1020/txtITOB-INVNR")
                )
                .Text = AssetMaster[i].New.OperationId;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(11);

            var l = new LoggingSystem();
            l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
            l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
            l.Type = "OPERATION ID";
            l.OldValue = AssetMaster[i].Old.OperationId;
            l.NewValue = AssetMaster[i].New.OperationId;
            AssetMaster[i].Old.OperationId = AssetMaster[i].New.OperationId;
            log.Add(l);
        }

        private static void eUpdateEquipmentDescription(int i)
        {
            ((GuiMainWindow)session.FindById("wnd[0]")).Maximize();
            ((GuiOkCodeField)session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nie02";
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).Text = AssetMaster[i].EquipmentNumber;
            ((GuiCTextField)session.FindById("wnd[0]/usr/ctxtRM63E-EQUNR")).CaretPosition = 8;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(0);
            ((GuiTextField)session.FindById(
                    "wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152B:SAPLITO0:1525/txtITOB-SHTXT")).Text =
                AssetMaster[i].New.EquipmentDescription;
            ((GuiTextField)session.FindById(
                "wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152B:SAPLITO0:1525/txtITOB-SHTXT")).CaretPosition = 24;
            ((GuiMainWindow)session.FindById("wnd[0]")).SendVKey(11);


            var l = new LoggingSystem();
            l.EquipmentNumber = AssetMaster[i].EquipmentNumber;
            l.Message = ((GuiStatusbar)session.FindById("wnd[0]/sbar")).Text;
            l.Type = "EQUIPMENT DESCRIPTION";
            l.OldValue = AssetMaster[i].Old.EquipmentDescription;
            l.NewValue = AssetMaster[i].New.EquipmentDescription;
            AssetMaster[i].Old.EquipmentDescription = AssetMaster[i].New.EquipmentDescription;
            log.Add(l);
        }

        private static void ActivateSAP()
        {
            //Get the Windows Running Object Table
            var sapROTWrapper = new CSapROTWrapper();
            //Get the ROT Entry for the SAP Gui to connect to the COM
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            //Get the reference to the Scripting Engine
            var engine = SapGuilRot.GetType()
                .InvokeMember("GetScriptingEngine", BindingFlags.InvokeMethod, null, SapGuilRot, null);
            //Get the reference to the running SAP Application Window
            var GuiApp = (GuiApplication)engine;
            //Get the reference to the first open connection
            var connection = (GuiConnection)GuiApp.Connections.ElementAt(0);
            //get the first available session
            session = (GuiSession)connection.Children.ElementAt(0);
            //Get the reference to the main "Frame" in which to send virtual key commands
            //GuiFrameWindow frame = (GuiFrameWindow)session.FindById("wnd[0]");
        }

        private static void OpenASM()
        {
            Process.Start(REPORT);
        }

        private static void WipeASM()
        {
            var fi = new FileInfo(REPORT);
            using (var excelPackage = new ExcelPackage(fi))
            {
                var myWorkbook = excelPackage.Workbook;
                var myWorksheet = myWorkbook.Worksheets[1];
                Console.WriteLine();
                myWorksheet.DeleteRow(2, myWorksheet.Dimension.End.Row);
                excelPackage.Save();
                Console.WriteLine();
                Console.WriteLine("File Wiped to " + myWorksheet.Dimension.End.Row + " Rows");
            }
        }

        private static void UpdateAssetMaster()
        {
            ReadFAR();
            ReadEQM();
            ProcessFAR();
            WriteToFile();
        }

        private static void WriteToFile()
        {
            WipeASM();
            var fi = new FileInfo(REPORT);
            using (var excelPackage = new ExcelPackage(fi))
            {
                var myWorkbook = excelPackage.Workbook;
                var myWorksheet = myWorkbook.Worksheets[1];
                Console.WriteLine();

                for (var i = 0; i < AssetMaster.Count; i++)
                {
                    myWorksheet.Cells[i + 2, 1].Formula = AssetMaster[i].AssetNumber;
                    myWorksheet.Cells[i + 2, 2].Formula = AssetMaster[i].EquipmentNumber;
                    myWorksheet.Cells[i + 2, 3].Value = AssetMaster[i].Old.AssetDescription;
                    myWorksheet.Cells[i + 2, 4].Value = AssetMaster[i].Old.AssetDescription;
                    myWorksheet.Cells[i + 2, 5].Value = AssetMaster[i].Old.EquipmentDescription;
                    myWorksheet.Cells[i + 2, 6].Value = AssetMaster[i].Old.EquipmentDescription;

                    myWorksheet.Cells[i + 2, 8].Value = AssetMaster[i].Old.OperationId;
                    myWorksheet.Cells[i + 2, 7].Value = AssetMaster[i].Old.OperationId;


                    myWorksheet.Cells[i + 2, 9].Value = AssetMaster[i].Old.SubTypeDescription;
                    myWorksheet.Cells[i + 2, 11].Value = AssetMaster[i].Old.SubTypeDescription;
                    myWorksheet.Cells[i + 2, 10].Value = AssetMaster[i].Old.SubType;
                    myWorksheet.Cells[i + 2, 12].Value = AssetMaster[i].Old.Weight;
                    myWorksheet.Cells[i + 2, 13].Value = AssetMaster[i].Old.WeightUnit;
                    myWorksheet.Cells[i + 2, 22].Value = AssetMaster[i].Old.ModelNumber;
                    myWorksheet.Cells[i + 2, 23].Value = AssetMaster[i].Old.ModelNumber;
                    myWorksheet.Cells[i + 2, 24].Value = AssetMaster[i].Old.SerialNumber;
                    myWorksheet.Cells[i + 2, 25].Value = AssetMaster[i].Old.SerialNumber;
                    myWorksheet.Cells[i + 2, 26].Value = AssetMaster[i].Old.AssetLocation;
                    myWorksheet.Cells[i + 2, 27].Value = AssetMaster[i].Old.EquipmentLocation;

                    if (String.IsNullOrEmpty(AssetMaster[i].New.EquipmentLocation))
                    {
                        myWorksheet.Cells[i + 2, 28].Value = AssetMaster[i].New.EquipmentLocation;
                    }
                    else
                    {
                        myWorksheet.Cells[i + 2, 28].Value = AssetMaster[i].Old.EquipmentLocation;
                    }



                    myWorksheet.Cells[i + 2, 29].Formula = AssetMaster[i].AcquisitionDate;
                    myWorksheet.Cells[i + 2, 30].Formula = AssetMaster[i].AcquisitionValue;
                    myWorksheet.Cells[i + 2, 31].Formula = AssetMaster[i].BookValue;

                    Console.Write("\r{0} Records Updated  ", i);
                }

                excelPackage.Save();
                Console.WriteLine();
                Console.WriteLine("File Saved");
            }
        }

        private static void ProcessFAR()
        {
            Console.WriteLine();
            var q = 1;
            for (var i = 0; i < AssetMaster.Count; i++)
            {
                var e = EquipmentMaster.FirstOrDefault(m => m.AssetNumber == AssetMaster[i].AssetNumber);
                if (e != null)
                {
                    AssetMaster[i].EquipmentNumber = e.EquipmentNumber;
                    AssetMaster[i].Old.EquipmentDescription = e.Old.EquipmentDescription;
                    AssetMaster[i].Old.OperationId = e.Old.OperationId;
                    AssetMaster[i].Old.SubType = e.Old.SubType;
                    AssetMaster[i].Old.SubTypeDescription = e.Old.SubTypeDescription;
                    AssetMaster[i].Old.Weight = e.Old.Weight;
                    AssetMaster[i].Old.WeightUnit = e.Old.WeightUnit;
                    AssetMaster[i].Old.Dimensions = e.Old.Dimensions;
                    AssetMaster[i].Old.ModelNumber = e.Old.ModelNumber;
                    AssetMaster[i].Old.SerialNumber = e.Old.SerialNumber;
                    AssetMaster[i].Old.AssetLocation = e.Old.AssetLocation;
                    AssetMaster[i].Old.EquipmentLocation = e.Old.EquipmentLocation;
                    AssetMaster[i].New.EquipmentLocation = e.Old.EquipmentLocation;
                    q++;
                    Console.Write("\r" + q + " Records Processed  " + (i - q) +
                                  " Equipments could not be found in EQM");
                }

            }
        }

        private static void ReadEQM()
        {
            var fi = new FileInfo(EQM);
            using (var excelPackage = new ExcelPackage(fi))
            {
                var myWorkbook = excelPackage.Workbook;
                var myWorksheet = myWorkbook.Worksheets[1];
                Console.WriteLine();
                for (var i = 1; i < myWorksheet.Dimension.End.Row; i++)
                {
                    var e = new Equipment();
                    e.EquipmentNumber = myWorksheet.Cells[i, 1].Text.Trim();
                    e.AssetNumber = myWorksheet.Cells[i, 2].Text.Trim();
                    e.Old.EquipmentDescription = myWorksheet.Cells[i, 3].Text.Trim();
                    e.Old.OperationId = myWorksheet.Cells[i, 4].Text.Trim();
                    e.Old.SubType = myWorksheet.Cells[i, 22].Text.Trim();
                    e.Old.SubTypeDescription = myWorksheet.Cells[i, 5].Text.Trim();
                    e.Old.Weight = myWorksheet.Cells[i, 6].Text.Trim();
                    e.Old.WeightUnit = myWorksheet.Cells[i, 7].Text.Trim();
                    e.Old.Dimensions = myWorksheet.Cells[i, 8].Text.Trim();
                    e.Old.ModelNumber = myWorksheet.Cells[i, 9].Text.Trim();
                    e.Old.SerialNumber = myWorksheet.Cells[i, 10].Text.Trim();
                    e.Old.EquipmentLocation = myWorksheet.Cells[i, 12].Text.Trim();

                    Console.Write("\r{0} Equipments Found   ", EquipmentMaster.Count);
                    EquipmentMaster.Add(e);
                }
            }
        }

        private static void ReadFAR()
        {
            var fi = new FileInfo(FAR);
            using (var excelPackage = new ExcelPackage(fi))
            {
                var myWorkbook = excelPackage.Workbook;
                var myWorksheet = myWorkbook.Worksheets[1];
                Console.WriteLine();
                for (var i = 2; i < myWorksheet.Dimension.End.Row; i++)
                {
                    var e = new Equipment();
                    e.AcquisitionValue = myWorksheet.Cells[i, 6].Text.Trim().Replace(",", string.Empty);
                    e.BookValue = myWorksheet.Cells[i, 8].Text.Trim().Replace(",", string.Empty);
                    e.AcquisitionDate = DateTime.Parse(DateTime.Parse(myWorksheet.Cells[i, 4].Value.ToString()).ToString("dd/MM/yyyy")).ToOADate().ToString();
                    e.AssetNumber = myWorksheet.Cells[i, 1].Text.Trim();
                    e.EquipmentNumber = myWorksheet.Cells[i, 3].Text.Trim();
                    e.Old.AssetDescription = myWorksheet.Cells[i, 5].Text.Trim();
                    e.Old.AssetLocation = myWorksheet.Cells[i, 10].Text.Trim();


                    Console.Write("\r{0} Assets Found   ", AssetMaster.Count);
                    if (e.AssetNumber != string.Empty && e.AssetNumber != "Asset") AssetMaster.Add(e);
                }
            }
        }
    }

}