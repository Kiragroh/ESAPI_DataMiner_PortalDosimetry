using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Data;
using VMS.CA.Scripting;
using VMS.DV.PD.Scripting;

namespace PDDataMining2
{

    static class Program
    {

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                using (Application app = Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }
        }

        static void Execute(Application app)
        {
            // very important line. otherwise it will not work. Read here for more info: https://www.reddit.com/r/esapi/comments/hkpa6q/pdsapi_problem_with_createtransientanalysis_in/
            VMS.DV.PD.UI.Base.VTransientImageDataMgr.CreateInstance(true);

            // Iterate through all patients
            int counter = 0;
           

            foreach (var patientSummary in app.PatientSummaries.Reverse())
            {
                
                DateTime startDate = new DateTime(2019, 01, 01);

                // Iterate through all patients

                #region useful methods for loop
                double GetMinXSize(PDBeam beam)
                {
                    double minX = 42;
                    foreach (ControlPoint cp in beam.Beam.ControlPoints)
                    {
                        if ((cp.JawPositions.X2 - cp.JawPositions.X1) / 10 < minX)
                        {
                            //jaw size returned in mm. div by 10 for cm.
                            minX = (cp.JawPositions.X2 - cp.JawPositions.X1) / 10;
                        }
                    }
                    return minX;
                }
                double GetMinYSize(PDBeam beam)
                {
                    double minY = 42;
                    foreach (ControlPoint cp in beam.Beam.ControlPoints)
                    {
                        if ((cp.JawPositions.Y2 - cp.JawPositions.Y1) / 10 < minY)
                        {
                            minY = (cp.JawPositions.Y2 - cp.JawPositions.Y1) / 10;
                        }
                    }
                    return minY;
                }
                double GetMaxXSize(PDBeam beam)
                {
                    double maxX = 0;
                    foreach (ControlPoint cp in beam.Beam.ControlPoints)
                    {
                        if ((cp.JawPositions.X2 - cp.JawPositions.X1) / 10 > maxX)
                        {
                            //jaw size returned in mm. div by 10 for cm.
                            maxX = (cp.JawPositions.X2 - cp.JawPositions.X1) / 10;
                        }
                    }
                    return maxX;
                }
                double GetMaxYSize(PDBeam beam)
                {
                    double maxY = 0;
                    foreach (ControlPoint cp in beam.Beam.ControlPoints)
                    {
                        if ((cp.JawPositions.Y2 - cp.JawPositions.Y1) / 10 > maxY)
                        {
                            maxY = (cp.JawPositions.Y2 - cp.JawPositions.Y1) / 10;
                        }
                    }
                    return maxY;
                }
                double GetAverageYSize(PDBeam beam)
                {
                    double averageY = 0;
                    foreach (ControlPoint cp in beam.Beam.ControlPoints)
                    {
                        
                            averageY += (cp.JawPositions.Y2 - cp.JawPositions.Y1) / 10;
                        
                    }
                    averageY = averageY / beam.Beam.ControlPoints.Count();
                    return averageY;
                }
                double GetAverageXSize(PDBeam beam)
                {
                    double averageX = 0;
                    foreach (ControlPoint cp in beam.Beam.ControlPoints)
                    {

                        averageX += (cp.JawPositions.X2 - cp.JawPositions.X1) / 10;

                    }
                    averageX = averageX / beam.Beam.ControlPoints.Count();
                    return averageX;
                }
                #endregion useful methods for loop

                // Retrieve patient information
                Patient p = app.OpenPatient(patientSummary);
                
                if (p != null)
                {
                    #region Data acquisition (sorry for using try/catch so much -> sometimes a mining process crashes because of one weird patient or field and for this Test-Mining-script I did not want this. Maybe will change it later.)

                    foreach (PDPlanSetup pdPlan in p.PDPlanSetups.OrderByDescending(y=>y.HistoryDateTime))
                {
                        
                        foreach (PDBeam pdBeam in pdPlan.Beams.Where(x => x.Beam.CreationDateTime > startDate))
                        {
                            counter++;

                            // Stop after when a few records have been found
                            if (counter > 10000000)
                                break;

                            
                            double gammaResult =0;
                            double maxDoseDifferenceResult=0;
                            double averageDoseDifferenceResult=0;
                            double maxDoseDifferenceRelResult = 0;
                            double maxDoseDifferenceRel2Result = 0;
                            double averageDoseDifferenceRelResult = 0;
                            double averageDoseDifferenceRel2Result = 0;

                            try
                            {
                                //
                                List<EvaluationTestDesc> evaluationTestDescs = new List<EvaluationTestDesc>();

                                EvaluationTestDesc evaluationTestDesc = new EvaluationTestDesc(EvaluationTestKind.MaxDoseDifferenceRelative, double.NaN, 0.95, false);
                                EvaluationTestDesc evaluationTestDesc2 = new EvaluationTestDesc(EvaluationTestKind.AverageDoseDifferenceRelative, double.NaN, 0.33, false);
                                                                
                                evaluationTestDescs.Add(evaluationTestDesc);
                                evaluationTestDescs.Add(evaluationTestDesc2);

                                PDTemplate pDTemplate = new PDTemplate(false, false, false, false, AnalysisMode.Relative, NormalizationMethod.MaxPredictedDose, false, 0.1, ROIType.CIAO, 0, 0.04, 4, false, evaluationTestDescs);

                                PortalDoseImage portaldoseImage = pdBeam.PortalDoseImages.LastOrDefault();

                                DoseImage predictedDoseImage = pdBeam.PredictedDoseImage;

                                PDAnalysis pDAnalysis = new PDAnalysis();

                                try
                                {
                                    pDAnalysis = portaldoseImage.CreateTransientAnalysis(pDTemplate, predictedDoseImage);
                                    EvaluationTest maxDoseDifferenceRel2Test = pDAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxDoseDifferenceRelative);
                                    maxDoseDifferenceRel2Result = Math.Round(maxDoseDifferenceRel2Test.TestValue * 100, 2);

                                    EvaluationTest averageDoseDifferenceRel2Test = pDAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.AverageDoseDifferenceRelative);
                                    averageDoseDifferenceRel2Result = Math.Round(averageDoseDifferenceRel2Test.TestValue * 100, 2);

                                }
                                catch { maxDoseDifferenceRel2Result=-1;
                                    averageDoseDifferenceRel2Result = -1;
                                }

                                //

                                PDAnalysis pdAnalysis = pdBeam.PortalDoseImages.LastOrDefault().Analyses.OrderBy(x => x.CreationDate).LastOrDefault();
                                if (pdAnalysis == null)
                                {
                                    int pdBeamcount = pdBeam.PortalDoseImages.Count();
                                    pdAnalysis = pdBeam.PortalDoseImages.FirstOrDefault().Analyses.OrderBy(x => x.CreationDate).LastOrDefault();
                                    if (pdAnalysis == null)
                                    {
                                        Console.WriteLine($"No Analysis for: {p.Id}, {pdPlan.Id}, {pdBeam.Id}");
                                        gammaResult = -2;
                                        maxDoseDifferenceResult = -2;
                                        averageDoseDifferenceResult = -2;
                                        maxDoseDifferenceRelResult = -2;
                                        averageDoseDifferenceRelResult = -2;
                                    }
                                    else
                                    {
                                        try
                                        {
                                            EvaluationTest gammaTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.GammaAreaLessThanOne);
                                            gammaResult = Math.Round(gammaTest.TestValue, 2);
                                        }
                                        catch { gammaResult = -3; }
                                        try
                                        {
                                            EvaluationTest maxDoseDifferenceTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxDoseDifference);
                                            if (pdBeam.Beam.ExternalBeam.SerialNumber.ToString() != "2479")
                                                maxDoseDifferenceResult = Math.Round(maxDoseDifferenceTest.TestValue * 100, 2);
                                            else
                                            {
                                                maxDoseDifferenceResult = Math.Round(maxDoseDifferenceTest.TestValue, 2);
                                            }
                                        }
                                        catch { maxDoseDifferenceResult = -3; }
                                        try
                                        {
                                            EvaluationTest averageDoseDifferenceTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.AverageDoseDifference);
                                            if (pdBeam.Beam.ExternalBeam.SerialNumber.ToString() != "2479")
                                                averageDoseDifferenceResult = Math.Round(averageDoseDifferenceTest.TestValue * 100, 2);
                                            else
                                            {
                                                averageDoseDifferenceResult = Math.Round(averageDoseDifferenceTest.TestValue, 2);
                                            }
                                        }
                                        catch { averageDoseDifferenceResult = -3; }
                                        try
                                        {
                                            EvaluationTest maxDoseDifferenceRelTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxDoseDifferenceRelative);
                                            maxDoseDifferenceRelResult = Math.Round(maxDoseDifferenceRelTest.TestValue*100, 2);
                                        }
                                        catch { maxDoseDifferenceRelResult = -3; }
                                        try
                                        {
                                            EvaluationTest averageDoseDifferenceRelTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.AverageDoseDifferenceRelative);
                                            averageDoseDifferenceRelResult = Math.Round(averageDoseDifferenceRelTest.TestValue*100, 2);
                                        }
                                        catch { averageDoseDifferenceRelResult = -3; }



                                        Console.WriteLine($"{p.Id}, {pdPlan.Id}, {pdBeam.Id},  {maxDoseDifferenceRel2Result},  {averageDoseDifferenceRel2Result}");
                                    }
                                }
                                else
                                {
                                    try { EvaluationTest gammaTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.GammaAreaLessThanOne);
                                        gammaResult = Math.Round(gammaTest.TestValue, 2);
                                    }
                                    catch { gammaResult = -3; }
                                    try { EvaluationTest maxDoseDifferenceTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxDoseDifference);
                                        if (pdBeam.Beam.ExternalBeam.SerialNumber.ToString() != "2479")
                                            maxDoseDifferenceResult = Math.Round(maxDoseDifferenceTest.TestValue * 100, 2);
                                        else
                                        {
                                            maxDoseDifferenceResult = Math.Round(maxDoseDifferenceTest.TestValue, 2);
                                        }
                                    }
                                    catch { maxDoseDifferenceResult = -3; }
                                    try { EvaluationTest averageDoseDifferenceTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.AverageDoseDifference);
                                        if (pdBeam.Beam.ExternalBeam.SerialNumber.ToString() != "2479")
                                            averageDoseDifferenceResult = Math.Round(averageDoseDifferenceTest.TestValue * 100, 2);
                                        else {
                                            averageDoseDifferenceResult = Math.Round(averageDoseDifferenceTest.TestValue, 2);
                                        }                                    }
                                    catch { averageDoseDifferenceResult = -3; }
                                    try
                                    {
                                        EvaluationTest maxDoseDifferenceRelTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxDoseDifferenceRelative);
                                        maxDoseDifferenceRelResult = Math.Round(maxDoseDifferenceRelTest.TestValue*100, 2);
                                    }
                                    catch { maxDoseDifferenceRelResult = -3; }
                                    try
                                    {
                                        EvaluationTest averageDoseDifferenceRelTest = pdAnalysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.AverageDoseDifferenceRelative);
                                        averageDoseDifferenceRelResult = Math.Round(averageDoseDifferenceRelTest.TestValue*100, 2);
                                    }
                                    catch { averageDoseDifferenceRelResult = -3; }



                                    Console.WriteLine($"{p.Id}, {pdPlan.Id}, {pdBeam.Id}, {maxDoseDifferenceRel2Result},  {averageDoseDifferenceRel2Result}");
                                
                                }
                            }
                            //catch { Console.WriteLine($"Error for: {p.Id}, {p.LastName}^{p.FirstName}, {pdPlans.Id}, {pdBeam.Id} "); }
                            catch {
                                Console.WriteLine($"Error for: {p.Id}, {pdPlan.Id}, {pdBeam.Id} ");
                                gammaResult = -1;
                                maxDoseDifferenceResult =-1;
                                averageDoseDifferenceResult = -1;
                                maxDoseDifferenceRelResult = -1;
                                averageDoseDifferenceRelResult = -1;
                            }

                        #region File-Writer -> User-Log-File-Syntax -> mainly copy from other project and therefore maybe strange names

                        string userLogPath;
                        StringBuilder userLogCsvContent = new StringBuilder();
                        if (Directory.Exists(@"\\Network-Path"))
                        {
                            userLogPath = @"\\variancom\daten\F?r Alle\Physik-Skripte\Output\PD-Mining\"+ System.DateTime.Now.ToString("yyyy-MM-dd") +"_PD-Mining.csv";
                        }
                        else
                        {
                            userLogPath = Path.GetTempFileName() + "_"+ System.DateTime.Now.ToString("yyyy-MM-dd") +"_PD - Mining.csv";
                        }


                        // add headers if the file doesn't exist
                        // list of target headers for desired dose stats
                        // in this case I want to display the headers every time so i can verify which target the distance is being measured for
                        // this is due to the inconsistency in target naming (PTV1/2 vs ptv45/79.2) -- these can be removed later when cleaning up the data
                        if (!File.Exists(userLogPath))
                        {
                            List<string> dataHeaderList = new List<string>();
                            dataHeaderList.Add("ID");
                            dataHeaderList.Add("Nachname");
                            dataHeaderList.Add("Vorname");
                            dataHeaderList.Add("Kurs");
                            dataHeaderList.Add("Plan");
                            dataHeaderList.Add("Beam");
                            dataHeaderList.Add("MU");
                                dataHeaderList.Add("Energy");
                                dataHeaderList.Add("DoseRate");
                                dataHeaderList.Add("BeamWeight");
                            dataHeaderList.Add("BeamCreation");
                            dataHeaderList.Add("Linac");
                            dataHeaderList.Add("MinX-Jaw");
                            dataHeaderList.Add("MinY-Jaw");
                                dataHeaderList.Add("AverageX-Jaw");
                                dataHeaderList.Add("AverageY-Jaw");
                                dataHeaderList.Add("AverageA-Jaws");
                            
                            dataHeaderList.Add("Gamma");
                            dataHeaderList.Add("MaxDoseDifference[KE]");
                            dataHeaderList.Add("AverageDoseDifference[KE]");
                            dataHeaderList.Add("MaxDoseDifferenceRelative[KE]");
                            dataHeaderList.Add("AverageDoseDifferenceRelative[KE]");
                                dataHeaderList.Add("ExtraTest-MaxDoseDiff[%]");
                                dataHeaderList.Add("ExtraTest-AverageDoseDiff[%]");


                                string concatDataHeader = string.Join(",", dataHeaderList.ToArray());

                            userLogCsvContent.AppendLine(concatDataHeader);
                        }
                       

                        List<object> userStatsList = new List<object>();
                        

                        userStatsList.Add(p.Id);
                        userStatsList.Add(p.LastName.Replace(",",""));
                        userStatsList.Add(p.FirstName.Replace(",", ""));
                        userStatsList.Add(pdPlan.PlanSetup.Course.Id.Replace(",", ""));
                        userStatsList.Add(pdPlan.Id.Replace(",", ""));
                        userStatsList.Add(pdBeam.Id.Replace(",", ""));
                        userStatsList.Add(Math.Round(pdBeam.PlannedMUs,0).ToString().Replace(",", "."));
                            userStatsList.Add(pdBeam.Beam.EnergyModeDisplayName.Replace(",", ""));
                            userStatsList.Add(pdBeam.Beam.DoseRate);
                            userStatsList.Add(Math.Round(pdBeam.Beam.WeightFactor, 2).ToString().Replace(",", "."));
                        userStatsList.Add(pdBeam.Beam.CreationDateTime);
                        userStatsList.Add(pdBeam.Beam.ExternalBeam.SerialNumber.ToString() == "1233" ? "Linac2" : (pdBeam.Beam.ExternalBeam.SerialNumber.ToString() == "1334" ? "Linac1" : (pdBeam.Beam.ExternalBeam.SerialNumber.ToString() == "2479"?"Linac3": pdBeam.Beam.ExternalBeam.SerialNumber.ToString())));
                        
                        userStatsList.Add(Math.Round(GetMinXSize(pdBeam), 2).ToString().Replace(",", "."));
                            userStatsList.Add(Math.Round(GetMinYSize(pdBeam), 2).ToString().Replace(",", "."));
                            userStatsList.Add(Math.Round(GetAverageXSize(pdBeam), 2).ToString().Replace(",", "."));
                            userStatsList.Add(Math.Round(GetAverageYSize(pdBeam), 2).ToString().Replace(",", "."));
                            userStatsList.Add(Math.Round(GetAverageXSize(pdBeam)* GetAverageYSize(pdBeam), 2).ToString().Replace(",", "."));
                        userStatsList.Add(gammaResult == -1 ? "Error" : (gammaResult == -2 ? "NoAnalysis" : (gammaResult.ToString().Replace(",", "."))));
                        userStatsList.Add(maxDoseDifferenceResult == -1 ? "Error" : (maxDoseDifferenceResult == -2 ? "NoAnalysis" : (maxDoseDifferenceResult == -3 ? "SpecificFail" : maxDoseDifferenceResult.ToString().Replace(",", "."))));
                        userStatsList.Add(averageDoseDifferenceResult == -1 ? "Error" : (averageDoseDifferenceResult == -2 ? "NoAnalysis" : (averageDoseDifferenceResult == -3 ? "SpecificFail" : averageDoseDifferenceResult.ToString().Replace(",", "."))));
                        userStatsList.Add(maxDoseDifferenceRelResult == -1 ? "Error" : (maxDoseDifferenceRelResult == -2 ? "NoAnalysis" : (maxDoseDifferenceRelResult == -3 ? "SpecificFail" : maxDoseDifferenceRelResult.ToString().Replace(",", "."))));
                        userStatsList.Add(averageDoseDifferenceRelResult == -1 ? "Error" : (averageDoseDifferenceRelResult == -2 ? "NoAnalysis" : (averageDoseDifferenceRelResult == -3 ? "SpecificFail" : averageDoseDifferenceRelResult.ToString().Replace(",", "."))));
                            userStatsList.Add(maxDoseDifferenceRel2Result == -1 ? "Error" : (maxDoseDifferenceRel2Result == -2 ? "NoAnalysis" : maxDoseDifferenceRel2Result.ToString().Replace(",", ".")));
                            userStatsList.Add(averageDoseDifferenceRel2Result == -1 ? "Error" : (averageDoseDifferenceRel2Result == -2 ? "NoAnalysis" : averageDoseDifferenceRel2Result.ToString().Replace(",", ".")));

                            // userStatsList.Add(planId);
                            //userStatsList.Add(course);
                            /*
                            var culture = new System.Globalization.CultureInfo("de-DE");
                            var day2 = culture.DateTimeFormat.GetDayName(System.DateTime.Today.DayOfWeek);
                            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                            string version = fvi.FileVersion;
                            string pc = Environment.MachineName.ToString();
                            string domain = Environment.UserDomainName.ToString();
                            string userId = Environment.UserName.ToString();
                            string scriptId = "exRay-Helper";
                            string date = System.DateTime.Now.ToString("yyyy-MM-dd");
                            string dayOfWeek = day2;
                            string time = string.Format("{0}:{1}", System.DateTime.Now.ToLocalTime().ToString("HH"), System.DateTime.Now.ToLocalTime().ToString("mm"));*/

                            string concatUserStats = string.Join(",", userStatsList.ToArray());

                        userLogCsvContent.AppendLine(concatUserStats);

                        File.AppendAllText(userLogPath, userLogCsvContent.ToString(),Encoding.Unicode);

                        #endregion
                    }

                }
                }
                #endregion Data acquisition


                // Close the current patient, otherwise we will not be able to open another patient
                app.ClosePatient();

            }
            // purpose: console stays open after finishing
            Console.WriteLine("DataMining finished?");
            Console.ReadLine();
            
        }

    }

}
