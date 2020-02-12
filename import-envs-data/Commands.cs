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
        [CommandMethod("IDA")]
        public void ImportEnvData()
        {

            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Local do arquivo de dados 
            // ************************************************************************s***************************************************
            var fileData = new StreamReader("Z:/Lisp/Arquivos_Teste/ida.csv");
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

                        sBlkName = sFileLine[0];
                        sBlkId = sFileLine[1];
                        ptBlkOrigin = new Point3d(Convert.ToDouble(sFileLine[2]), Convert.ToDouble(sFileLine[3]), 0);
                        //sBlkCoord = sFileLine[2].Split(',');
                        //ptBlkOrigin = new Point3d(Convert.ToDouble(sBlkCoord[0]), Convert.ToDouble(sBlkCoord[1]), 0);
                        sBlkRot = sFileLine[4];

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
                                    rb.Add(new TypedValue((int)DxfCode.ExtendedDataHandle, eBlk.Handle));

                                    xRec.Data = rb;

                                    dbExt.SetAt(sBlkId, xRec);
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

        [CommandMethod("LET")]
        public void ConnectConduits()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Local do arquivo de dados 
            // ***************************************************************************************************************************
            var fileData = new StreamReader("Z:/Lisp/Arquivos_Teste/let.csv");
            // ***************************************************************************************************************************

            string[] sLine;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                BlockTable BlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                BlockTableRecord BlkTblRec = tr.GetObject(BlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                DBObject dbModelSpace = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                ObjectId extId = dbModelSpace.ExtensionDictionary;
                DBDictionary dbExt = (DBDictionary)tr.GetObject(extId, OpenMode.ForRead);
                Point2d ptVert;
                int indexVert;

                while (!fileData.EndOfStream)
                {
                    sLine = fileData.ReadLine().Split(';');
                    indexVert = 0;

                    using(Polyline oPLine = new Polyline())
                    {
                        foreach(string address in sLine)
                        {
                            if(address != "")
                            {
                                if (address.Contains(','))
                                {
                                    string[] coords = address.Split(',');
                                    ptVert = new Point2d(Convert.ToDouble(coords[0]), Convert.ToDouble(coords[1]));
                                }
                                else
                                {
                                    ptVert = GetPtFromHandleBlock(db, dbExt, address);
                                }

                                //oPLine.SetPointAt(indexPoint, ptVert);
                                oPLine.AddVertexAt(indexVert, ptVert, 0, 0, 0);
                                indexVert++;
                            }
                        }
                        BlkTblRec.AppendEntity(oPLine);
                        tr.AddNewlyCreatedDBObject(oPLine, true);
                    }
                }
                tr.Commit();
            }
        }

        public Point2d GetPtFromHandleBlock(Database db, DBDictionary dbExt, string idHn)
        {
            BlockReference blk;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Pega o handle a partir do id dado 
                ObjectId idId = dbExt.GetAt(idHn);
                Xrecord xRec = (Xrecord)tr.GetObject(idId, OpenMode.ForRead);
                ResultBuffer rb = xRec.Data;
                TypedValue[] tp = rb.AsArray();
                string hand = tp[0].Value as string;

                long ln = Convert.ToInt64(hand, 16);
                Handle hn = new Handle(ln);
                ObjectId id = db.GetObjectId(false, hn, 0);

                blk = (BlockReference)tr.GetObject(id, OpenMode.ForRead);

            }

            return new Point2d(blk.Position.X, blk.Position.Y);
        }
    }
}
