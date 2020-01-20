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
        public void Test()
        {

            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            var fileData = new StreamReader("C:/Users/weiglas.ribeiro.LANGAMER/Desktop/teste.csv");
        
            String blkPath = "Z:/Lisp/BLOCOS";
            String[] line;
            String blkName;
            String[] blkCoord;
            Point3d ptBlkOrigin;
            String blkRot;

            using (var tr = db.TransactionManager.StartTransaction())
            {

                BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                ObjectId blkRecId = ObjectId.Null;

                using (var acBlkTblRec = new BlockTableRecord())
                {
                    while (!fileData.EndOfStream)
                    {
                        line = fileData.ReadLine().Split(';');

                        blkName = line[0];
                        blkCoord = line[1].Split(',');
                        ptBlkOrigin = new Point3d(Convert.ToDouble(blkCoord[0]), Convert.ToDouble(blkCoord[1]), 0);
                        blkRot = line[2];

                        if (!acBlkTbl.Has(blkName))
                        {
                            try
                            {
                                using (var blkDb = new Database(false, true))
                                {
                                    blkDb.ReadDwgFile(blkPath + "/" + blkName + ".dwg", FileOpenMode.OpenForReadAndAllShare, true, "");
                                    db.Insert(blkName, blkDb, true);
                                }
                            }
                            catch
                            {

                            }
                        }
                       
                        blkRecId = acBlkTbl[blkName];    

                        if (blkRecId != ObjectId.Null)
                        {
                            using (var acBlkRef = new BlockReference(ptBlkOrigin, blkRecId))
                            {
                                BlockTableRecord acCurSpaceBlkTblRec;
                                acCurSpaceBlkTblRec = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                                tr.AddNewlyCreatedDBObject(acBlkRef, true);
                            }
                        }


                    }

                }

                fileData.Close();
                tr.Commit();
            }
        }
    }
}
