
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace BlockSyncAuto
{
    public class BlockSyncAuto
    {
        private const string SyncLayer = "YHK-HOLE";
        private static string SourceDocName;
        private static string TargetDocName;

        [CommandMethod("SETBLOCKSYNC")]
        public void SetBlockSync()
        {
            var docs = Application.DocumentManager;
            var ed = docs.MdiActiveDocument.Editor;

            SourceDocName = docs.MdiActiveDocument.Name;
            ed.WriteMessage($"\n已设置源图纸: {SourceDocName}");

            PromptStringOptions pso = new PromptStringOptions("\n请输入目标图纸名（如2.dwg）：");
            PromptResult pr = ed.GetString(pso);
            if (pr.Status != PromptStatus.OK) return;
            TargetDocName = pr.StringResult;

            bool found = false;
            foreach (Document doc in docs)
            {
                if (doc.Name.ToLower().EndsWith(TargetDocName.ToLower()))
                {
                    TargetDocName = doc.Name;
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                ed.WriteMessage($"\n没有找到目标图纸: {TargetDocName}");
                return;
            }

            docs.DocumentActivated -= OnActivated;
            docs.DocumentActivated += OnActivated;
            ed.WriteMessage("\n已开启双向自动同步。");
        }

        [CommandMethod("CLEARBLOCKSYNC")]
        public void ClearBlockSync()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(doc.Database),
                    OpenMode.ForWrite);

                int count = 0;
                foreach (ObjectId id in modelSpace)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                    if (ent != null && ent.Layer == SyncLayer)
                    {
                        ent.Erase();
                        count++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n已清除 {count} 个同步图层上的实体。");
            }
        }

        private static void OnActivated(object sender, DocumentCollectionEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(SourceDocName) || string.IsNullOrEmpty(TargetDocName)) return;

                var activeDoc = Application.DocumentManager.MdiActiveDocument;
                if (activeDoc == null) return;

                var ed = activeDoc.Editor;

                if (activeDoc.Name == TargetDocName || activeDoc.Name == SourceDocName)
                {
                    Document otherDoc = null;
                    foreach (Document doc in Application.DocumentManager)
                    {
                        if (doc.Name == (activeDoc.Name == SourceDocName ? TargetDocName : SourceDocName))
                        {
                            otherDoc = doc;
                            break;
                        }
                    }

                    if (otherDoc == null)
                    {
                        ed.WriteMessage($"\n未找到对应图纸");
                        return;
                    }

                    Document srcDoc = activeDoc.Name == SourceDocName ? activeDoc : otherDoc;
                    Document dstDoc = activeDoc.Name == SourceDocName ? otherDoc : activeDoc;

                    using (DocumentLock lockSrc = srcDoc.LockDocument())
                    using (DocumentLock lockDst = dstDoc.LockDocument())
                    {
                        SyncBlocks(srcDoc, dstDoc, ed);
                    }
                }
            }
            catch (System.Exception ex)
            {
                var ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                if (ed != null)
                {
                    ed.WriteMessage($"\n同步异常: {ex.Message}");
                    ed.WriteMessage($"\n异常详情: {ex.StackTrace}");
                }
            }
        }

        private static void EnsureLayerExists(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord();
                ltr.Name = layerName;
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        private static void SyncBlocks(Document srcDoc, Document dstDoc, Editor ed)
        {
            try
            {
                // 读取源文件中的块参照
                List<BlockData> sourceBlocks = new List<BlockData>();
                using (Transaction tr = srcDoc.Database.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(srcDoc.Database.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId id in ms)
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent is BlockReference br && ent.Layer == SyncLayer)
                        {
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            sourceBlocks.Add(new BlockData
                            {
                                BlockName = btr.Name,
                                Position = br.Position,
                                XScale = br.ScaleFactors.X,
                                YScale = br.ScaleFactors.Y,
                                ZScale = br.ScaleFactors.Z,
                                Rotation = br.Rotation
                            });
                        }
                    }
                    tr.Commit();
                }

                if (sourceBlocks.Count == 0)
                {
                    ed.WriteMessage($"\n源文件中没有需要同步的块。");
                    return;
                }

                // 在目标文件中同步
                using (Transaction tr = dstDoc.Database.TransactionManager.StartTransaction())
                {
                    // 确保图层存在
                    EnsureLayerExists(dstDoc.Database, tr, SyncLayer);

                    // 获取目标文件的模型空间
                    BlockTable bt = (BlockTable)tr.GetObject(dstDoc.Database.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // 删除目标文件中的现有块
                    List<ObjectId> toDelete = new List<ObjectId>();
                    foreach (ObjectId id in ms)
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent != null && ent.Layer == SyncLayer)
                        {
                            toDelete.Add(id);
                        }
                    }

                    foreach (ObjectId id in toDelete)
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                        ent.Erase();
                    }

                    // 导入块定义并创建新的块参照
                    foreach (var blockData in sourceBlocks)
                    {
                        if (!bt.Has(blockData.BlockName))
                        {
                            // 从源文件获取块定义
                            ObjectId srcBlockId;
                            using (Transaction trSrc = srcDoc.Database.TransactionManager.StartTransaction())
                            {
                                BlockTable btSrc = (BlockTable)trSrc.GetObject(srcDoc.Database.BlockTableId, OpenMode.ForRead);
                                srcBlockId = btSrc[blockData.BlockName];
                                trSrc.Commit();
                            }

                            // 克隆块定义到目标文件
                            ObjectIdCollection ids = new ObjectIdCollection();
                            ids.Add(srcBlockId);
                            IdMapping mapping = new IdMapping();
                            srcDoc.Database.WblockCloneObjects(ids, dstDoc.Database.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
                        }

                        // 创建新的块参照
                        BlockReference newBr = new BlockReference(blockData.Position, bt[blockData.BlockName]);
                        newBr.Layer = SyncLayer;
                        newBr.ScaleFactors = new Scale3d(blockData.XScale, blockData.YScale, blockData.ZScale);
                        newBr.Rotation = blockData.Rotation;
                        ms.AppendEntity(newBr);
                        tr.AddNewlyCreatedDBObject(newBr, true);
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n成功同步 {sourceBlocks.Count} 个块参照。");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n同步出错: {ex.Message}");
                ed.WriteMessage($"\n错误详情: {ex.StackTrace}");
                throw;
            }
        }
    }
}
