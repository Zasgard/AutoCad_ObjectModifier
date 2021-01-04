using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.EditorInput;


namespace RevA
{
    public class Class1
    {
        private Document Open(string address)
        {
            string strFileName = address;
            Document curDoc;
            DocumentCollection acDocMgr = Application.DocumentManager;
            if (File.Exists(strFileName)) 
            { 
                curDoc=acDocMgr.Open(strFileName, false);
            } 
            else 
            { 
                acDocMgr.MdiActiveDocument.Editor.WriteMessage("File " + strFileName + " does not exist.");
                curDoc = acDocMgr.MdiActiveDocument;
            }
            return curDoc;
        }

        private static void modifyAtt(Database db)
        {
            // use an OpenCloseTransaction which is not related to a document or database
            using (var tr = new OpenCloseTransaction())
            {

                // get the database block table
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                // if the block table contains a block definitions named Atronix_tb_text
                if (bt.Has("ATRONIX_TB_TEXT"))
                {

                    // open the block definition
                    var btr = (BlockTableRecord)tr.GetObject(bt["ATRONIX_TB_TEXT"], OpenMode.ForRead);

                    // get the inserted block references ObjectIds
                    var ids = btr.GetBlockReferenceIds(true, true);

                    // if any...
                    for (int temp = 0; temp < ids.Count; temp++)
                    {
                        // open the first block reference
                        var br = (BlockReference)tr.GetObject(ids[temp], OpenMode.ForRead);

                        // iterate through the attribute collection
                        foreach (ObjectId id in br.AttributeCollection)
                        {
                            DBObject dbObj = tr.GetObject(id, OpenMode.ForRead) as DBObject;

                            if(dbObj is AttributeReference)
                            {

                                tr.GetObject(id, OpenMode.ForRead);
                                AttributeReference att = dbObj as AttributeReference;

                                // open the attribute reference
                                //var attRef = (AttributeReference)tr.GetObject(id, OpenMode.ForWrite);

                                // if the attribute tag is equal to 'tag'
                                if (att.Tag.Equals("REV_NUM", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "A";
                                    //
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV1_NUM", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "A";
                                    //att.AdjustAlignment(db);
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV1_BY", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "JL";
                                    //att.AdjustAlignment(db);
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV1_DATE", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "12/16/2020";
                                    //att.AdjustAlignment(db);
                                    att.DowngradeOpen();
                                }

                                else if (att.Tag.Equals("REV1_DESC", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "AS BUILT";
                                    //att.AdjustAlignment(db);
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV1_DESC_1", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "AS BUILT";
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV1_DESC_2", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "";
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV2_NUM", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "";
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV2_BY", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "";
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV2_DATE", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "";
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV2_DESC", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "";
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV2_DESC_1", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "";
                                    att.DowngradeOpen();
                                }
                                else if (att.Tag.Equals("REV2_DESC_2", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    att.UpgradeOpen();
                                    att.TextString = "";
                                    att.DowngradeOpen();
                                }
                            }
                        }
                    }
                }

                tr.Commit();
            }
        }

        private static void RevChangeA(Document d)
        {

            // get the documents collection
            var doc = d;

            //lock the document
            using (doc.LockDocument())
            {
                // get the document database
                var db = doc.Database;
                modifyAtt(db);
            }
        }

        static private List<string> SelectFiles()
        {
            var result = new List<string>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            OpenFileDialog ofd = new OpenFileDialog("Select the files to change revision", null, "dwg", null, OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK)
            {
                return result;
            }
            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\users\jlee\Documents\track.txt"))
            
            foreach(string line in ofd.GetFilenames())
            {
                result.Add(line);
            }
            return result;
            //ed.WriteMessage("\nFile selected is \"{0}\".", ofd.Filename);
        }

        [CommandMethod ("ra")]
        public void test1()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            RevChangeA(doc);
        }
        
        [CommandMethod ("ras", CommandFlags.Session)]   //doesn't work do not use
        public void test2()
        {
            var files = SelectFiles();
            var aoDocMgr = Application.DocumentManager;
            foreach(string dwg in files)
            {
                Document doc = aoDocMgr.Open(dwg, false);
                aoDocMgr.MdiActiveDocument = doc;
                using (doc.LockDocument())
                {
                    var db = doc.Database;
                    modifyAtt(db);
                    doc.Database.SaveAs(dwg, DwgVersion.Current);
                }
                doc.CloseAndDiscard();
            }
        }
        
        [CommandMethod("REVA")]
        public void dbRevChange()
        {
            var files = SelectFiles();
            foreach(string dwg in files)
            {
                using (Database db = new Database())
                {
                    db.ReadDwgFile(dwg, FileOpenMode.OpenForReadAndAllShare, false, null);
                    //Closing input makes sure the whole dwg is read from disk and it also closes the file so you can save as the same name
                    db.CloseInput(true);

                    //store current working database to reset
                    Database currDb = HostApplicationServices.WorkingDatabase;

                    //set working database
                    HostApplicationServices.WorkingDatabase = db;

                    modifyAtt(db);

                    db.SaveAs(dwg, true, DwgVersion.Current, db.SecurityParameters);

                    //resetting the database
                    HostApplicationServices.WorkingDatabase = currDb;
                }
            }
            Application.ShowAlertDialog("Files have been Processed");
        }
        //Class Structure ends here
    }
}
