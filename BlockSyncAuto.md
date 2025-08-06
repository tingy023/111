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
                BlockTableRecord ms = tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(doc.Database),
                    OpenMode.ForWrite) as BlockTableRecord;

                int count = 0;
                List<ObjectId> toDelete = new List<ObjectId>();
                
                foreach (ObjectId id in ms)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null && ent.Layer == SyncLayer)
                    {
                        toDelete.Add(id);
                        count++;
                    }
                }

                foreach (ObjectId id in toDelete)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                    ent.Erase();
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

                    // 先锁定目标文档
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
            LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
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
                // 临时数据库用于中转
                using (Database tempDb = new Database(true, true))
                {
                    // 第一步：收集源文档中的块到临时数据库
                    ObjectIdCollection sourceIds = new ObjectIdCollection();
                    using (Transaction tr = srcDoc.Database.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord ms = tr.GetObject(
                            SymbolUtilityServices.GetBlockModelSpaceId(srcDoc.Database),
                            OpenMode.ForRead) as BlockTableRecord;

                        foreach (ObjectId id in ms)
                        {
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent is BlockReference && ent.Layer == SyncLayer)
                            {
                                sourceIds.Add(id);
                            }
                        }
                        tr.Commit();
                    }

                    if (sourceIds.Count == 0)
                    {
                        ed.WriteMessage($"\n源文件中没有需要同步的块。");
                        return;
                    }

                    // 第二步：清理目标文档中的旧块
                    using (Transaction tr = dstDoc.Database.TransactionManager.StartTransaction())
                    {
                        EnsureLayerExists(dstDoc.Database, tr, SyncLayer);

                        BlockTableRecord ms = tr.GetObject(
                            SymbolUtilityServices.GetBlockModelSpaceId(dstDoc.Database),
                            OpenMode.ForWrite) as BlockTableRecord;

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

                        tr.Commit();
                    }

                    // 第三步：通过临时数据库克隆块到目标文档
                    using (Transaction tr = dstDoc.Database.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord ms = tr.GetObject(
                            SymbolUtilityServices.GetBlockModelSpaceId(dstDoc.Database),
                            OpenMode.ForWrite) as BlockTableRecord;

                        // 使用临时映射确保所有依赖项都被正确克隆
                        IdMapping mapping = new IdMapping();
                        srcDoc.Database.WblockCloneObjects(sourceIds, ms.ObjectId, mapping, DuplicateRecordCloning.Replace, false);

                        tr.Commit();
                        ed.WriteMessage($"\n成功同步 {sourceIds.Count} 个块。");
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n同步出错: {ex.Message}");
                ed.WriteMessage($"\n错误详情: {ex.StackTrace}");
            }
        }
    }
}
