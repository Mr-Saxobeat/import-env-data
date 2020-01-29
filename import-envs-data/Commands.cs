using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AcadPlugin
{
    public class Commands
    {
        [CommandMethod("IMPORTENVDATA")]
        public void ImportEnvData()
        {

            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Local do arquivo de dados 
            // ***************************************************************************************************************************
            var fileData = new StreamReader("Z:/Lisp/Arquivos_Teste/test-import-env-data.csv");
            // ***************************************************************************************************************************


            // Local dos blocos a serem carregados
            // ***************************************************************************************************************************
            String blkPath = "Z:/Lisp/BLOCOS";
            // ***************************************************************************************************************************

            string[] sFileLine;
            string sBlkId;
            string sBlkName;
            string[] sBlkCoord;
            Point3d ptBlkOrigin;
            string sBlkRot;

            using (var tr = db.TransactionManager.StartTransaction())
            {

                BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                ObjectId blkRecId = ObjectId.Null;

                using (var acBlkTblRec = new BlockTableRecord())
                {
                    DBObject dbModelSpace = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                    while (!fileData.EndOfStream)
                    {
                        sFileLine = fileData.ReadLine().Split(';');

                        sBlkId = sFileLine[0];
                        sBlkName = sFileLine[1];
                        sBlkCoord = sFileLine[2].Split(',');
                        ptBlkOrigin = new Point3d(Convert.ToDouble(sBlkCoord[0]), Convert.ToDouble(sBlkCoord[1]), 0);
                        sBlkRot = sFileLine[3];

                        if (!acBlkTbl.Has(sBlkName))
                        {
                            try
                            {
                                using (var blkDb = new Database(false, true))
                                {
                                    blkDb.ReadDwgFile(blkPath + "/" + sBlkName + ".dwg", FileOpenMode.OpenForReadAndAllShare, true, "");
                                    ObjectId blkId = db.Insert(sBlkName, blkDb, true);
                                }
                            }
                            catch
                            {

                            }
                        }
                       
                        blkRecId = acBlkTbl[sBlkName];    

                        if (blkRecId != ObjectId.Null)
                        {
                            using (var acBlkRef = new BlockReference(ptBlkOrigin, blkRecId))
                            {
                                BlockTableRecord acCurSpaceBlkTblRec;
                                acCurSpaceBlkTblRec = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                                tr.AddNewlyCreatedDBObject(acBlkRef, true);

                                Entity eBlk = (Entity)tr.GetObject(acBlkRef.Id, OpenMode.ForRead);

                                ObjectId extId = dbModelSpace.ExtensionDictionary;

                                if (extId == ObjectId.Null)
                                {
                                    dbModelSpace.CreateExtensionDictionary();
                                    extId = dbModelSpace.ExtensionDictionary;
                                }

                                DBDictionary dbExt = (DBDictionary)tr.GetObject(extId, OpenMode.ForWrite);

                                //if (!dbExt.Contains("Ids"))
                                //{
                                    Xrecord xRec = new Xrecord();
                                    ResultBuffer rb = new ResultBuffer();
                                    rb.Add(new TypedValue((int)DxfCode.ExtendedDataReal, Convert.ToDouble(sBlkId)));
                                    rb.Add(new TypedValue((int)DxfCode.ExtendedDataHandle, eBlk.Handle));

                                    xRec.Data = rb;

                                    dbExt.SetAt("Ids", xRec);
                                    tr.AddNewlyCreatedDBObject(xRec, true);
                                //}
                            }
                        }

                        


                    }

                }

                fileData.Close();
                tr.Commit();
            }
        }

        [CommandMethod("LIGAELETRODUTOS")]
        public void ConnectConduits()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Local do arquivo de dados 
            // ***************************************************************************************************************************
            var fileData = new StreamReader("Z:/Lisp/Arquivos_Teste/test-liga-eletro.csv");
            // ***************************************************************************************************************************

            string[] sLine;
            //String[] sCondCoord;
            string[] sIds;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                DBObject dbModelSpace = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                ObjectId extId = dbModelSpace.ExtensionDictionary;
                DBDictionary dbExt = (DBDictionary)tr.GetObject(extId, OpenMode.ForRead);
                ObjectId testIds = dbExt.GetAt("Ids");
                Xrecord xRec = (Xrecord)tr.GetObject(testIds, OpenMode.ForRead);
                

                while (!fileData.EndOfStream)
                {
                    sLine = fileData.ReadLine().Split(';');

                    sIds = new String[2];
                    sIds[0] = sLine[0];
                    sIds[1] = sLine[1];

                    //var oLine = new Line(new Point3d(Convert.ToDouble(sCondCoord[0]), Convert.ToDouble(sCondCoord[1]), 0));
                }
            }
        }
    }
}
