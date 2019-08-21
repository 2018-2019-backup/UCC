using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dObjects;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance;
using System.IO;
using System.Diagnostics;
using System.Data;
using Autodesk.ProcessPower.DataObjects;

namespace PID
{
    public class Command
    {
        private List<KeyValuePair<string, string>> list_Properties;
        private List<PnPProjectDrawing> oDwgList;
        private Project prj;
        private List<string> selected = new List<string>();

        [CommandMethod("Partnumber", CommandFlags.Session)]

        public void Partnumber()
        {
            PlantProject mainPrj = PlantApplication.CurrentProject;
            prj = mainPrj.ProjectParts["Piping"];
            oDwgList = prj.GetPnPDrawingFiles();

            if (oDwgList.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("3D Drawing not available in the Project");
                return;
            }

            List<string> listvalue = new List<string>();
            for (int n = 0; n < oDwgList.Count; n++)
            {
                listvalue.Add(System.IO.Path.GetFileNameWithoutExtension(oDwgList[n].ResolvedFilePath));
            }

            double fulllengthuser = 0;


            System.Windows.Forms.DialogResult dr = System.Windows.Forms.MessageBox.Show("Click 'Yes' if Full length is 18' or Click 'No' if Full length is 17'", "Partnumber", System.Windows.Forms.MessageBoxButtons.YesNo);

            if (dr == System.Windows.Forms.DialogResult.Yes)
                fulllengthuser = 18 * 12;
            else if (dr == System.Windows.Forms.DialogResult.No)
                fulllengthuser = 17 * 12;
            else
                return;

            DataLinksManager dlm = prj.DataLinksManager;// DataLinksManager.GetManager(db);
                                                        //Application.DocumentManager.CurrentDocument.CloseAndDiscard();

            //for (int n = 0; n < oDwgList.Count; n++)
            //{
            //string filename = System.IO.Path.GetFileNameWithoutExtension(oDwgList[n].ResolvedFilePath);

            //Document mcdoc = Application.DocumentManager.Open(oDwgList[n].ResolvedFilePath, false);
            Document mcdoc = Application.DocumentManager.MdiActiveDocument;
            //Application.DocumentManager.MdiActiveDocument = mcdoc;

            Editor ed = mcdoc.Editor;
            Database db = mcdoc.Database;
            using (DocumentLock Loc = mcdoc.LockDocument())
            {
                using (Transaction tr = mcdoc.TransactionManager.StartTransaction())
                {
                    PromptSelectionResult psr = ed.SelectAll();

                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    List<PartSizeProperties> psp = new List<PartSizeProperties>();
                    List<KeyValuePair<string, string>> prop = new List<KeyValuePair<string, string>>();

                    foreach (ObjectId objid in btr)
                    {
                        try
                        {
                            Entity en = tr.GetObject(objid, OpenMode.ForWrite) as Entity;
                            if (en != null)
                            {

                                Debug.WriteLine(en.ToString());
                                list_Properties = dlm.GetAllProperties(objid, true);
                                string partnumber = string.Empty;
                                string PartCategory = getvaluefrompropertylist("PnPClassName");

                                if (PartCategory.StartsWith("pipe", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    double NominalDiameter = Convert.ToDouble(getvaluefrompropertylist("Size").Replace("\"", ""));
                                    double CutLength = Convert.ToDouble(getvaluefrompropertylist("CutLength"));
                                    string prefix = "C-59997-005-048-";// getvaluefrompropertylist("UCC_PartNumber_Prefix");

                                    if (CutLength == fulllengthuser)
                                    {
                                        if (NominalDiameter == 4)
                                        {
                                            partnumber = "174001";
                                        }
                                        else if (NominalDiameter == 5)
                                        {
                                            partnumber = "174002";
                                        }
                                        else if (NominalDiameter == 6)
                                        {
                                            partnumber = "174003";
                                        }
                                        else if (NominalDiameter == 7)
                                        {
                                            partnumber = "174004";
                                        }
                                        else if (NominalDiameter == 8)
                                        {
                                            partnumber = "174005";
                                        }
                                        else if (NominalDiameter == 9)
                                        {
                                            partnumber = "174006";
                                        }
                                        else if (NominalDiameter == 10)
                                        {
                                            partnumber = "174007";
                                        }
                                        else if (NominalDiameter == 11)
                                        {
                                            partnumber = "174012";
                                        }
                                        else if (NominalDiameter == 12)
                                        {
                                            partnumber = "174008";
                                        }
                                        else if (NominalDiameter == 14)
                                        {
                                            partnumber = "174009";
                                        }
                                    }
                                    else if (CutLength > fulllengthuser)
                                    {
                                        partnumber = "Cut Length has exceeded Full Length";
                                    }
                                    else
                                    {
                                        partnumber = prefix + NominalDiameter.ToString().PadLeft(2, '0') + "-";

                                        string feet = ((int)CutLength / 12).ToString().PadLeft(2, '0') + "-";

                                        string inch = ((int)(CutLength % 12)).ToString().PadLeft(2, '0') + "-";

                                        //string decimalval = ((int)(((CutLength % 12) % 1) * 16) / 1).ToString().PadLeft(2, '0');
                                        string decimalval;
                                        if (Math.Round((((CutLength % 12) % 1) * 16)) < 16)
                                        {
                                            decimalval = Math.Round((((CutLength % 12) % 1) * 16)).ToString().PadLeft(2, '0');
                                        }
                                        else
                                        {
                                            inch = ((int)(CutLength % 12) + 1).ToString().PadLeft(2, '0') + "-";
                                            decimalval = "00";
                                        }
                                        partnumber = partnumber + feet + inch + decimalval;
                                    }


                                    updatelist("UCCPARTNO", partnumber);

                                    int rowId1 = dlm.FindAcPpRowId(objid);
                                    StringCollection strNames = new StringCollection();
                                    StringCollection strValues = new StringCollection();
                                    for (int i = 0; i < list_Properties.Count; i++)
                                    {
                                        String name = list_Properties[i].Key;
                                        String value = list_Properties[i].Value;
                                        strNames.Add(name);
                                        strValues.Add(value);
                                    }

                                    dlm.SetProperties(rowId1, strNames, strValues);
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    tr.Commit();
                }
            }
            //mcdoc.CloseAndSave(mcdoc.Name);
            //}
        }
        public string getvaluefrompropertylist(string property_name)
        {
            string returnpropery = "";
            if (list_Properties.Where(a => a.Key.Equals(property_name)).Count() > 0)
            {
                returnpropery = list_Properties.Where(a => a.Key.Equals(property_name)).FirstOrDefault().Value;
            }
            return returnpropery;
        }

        public void updatelist(string key, string value)
        {
            KeyValuePair<string, string> Toremove = list_Properties.FirstOrDefault(v => v.Key.Equals(key));
            KeyValuePair<string, string> Toadd = new KeyValuePair<string, string>(key, value);
            list_Properties.Remove(Toremove);
            list_Properties.Add(Toadd);
        }
    }
}
