namespace mpPrPosition
{
    using System;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Runtime;
    using Autodesk.AutoCAD.Windows;
    using ModPlus;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using mpProductInt;
    using Exception = System.Exception;

    public class ObjectContextMenu
    {
        private const string LangItem = "mpPrPosition";
        public static ContextMenuExtension MpPrPositionMarkCme;

        public static void Attach()
        {
            if (MpPrPositionMarkCme == null)
            {
                MpPrPositionMarkCme = new ContextMenuExtension();
                var miEnt = new MenuItem(Language.GetItem(LangItem, "h9"));
                miEnt.Click += SendCommand;
                MpPrPositionMarkCme.MenuItems.Add(miEnt);
            }

            var rxcEnt = RXObject.GetClass(typeof(Entity));
            Application.AddObjectContextMenuExtension(rxcEnt, MpPrPositionMarkCme);
        }

        public static void Detach()
        {
            if (MpPrPositionMarkCme != null)
            {
                var rxcEnt = RXObject.GetClass(typeof(Entity));
                Application.RemoveObjectContextMenuExtension(rxcEnt, MpPrPositionMarkCme);
            }
        }

        public static void SendCommand(object o, EventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.SendStringToExecute(
                "_.MPLEADERWITHMARKFORPRODUCT ", true, false, false);
        }

        [CommandMethod("ModPlus", "mpLeaderWithMarkForProduct", CommandFlags.UsePickSet)]
        public static void AddLeaderWithMark()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            var res = ed.SelectImplied();
            if (res.Status != PromptStatus.OK)
                return;
            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        foreach (var objectId in res.Value.GetObjectIds())
                        {
                            var ent = (Entity)tr.GetObject(objectId, OpenMode.ForWrite);

                            // Расширенные данные в виде спец.класса
                            var productInsert = (MpProductToSave)XDataHelpersForProducts.NewFromEntity(ent);
                            if (productInsert == null)
                                continue;

                            // Теперь получаем сам продукт
                            var product = MpProduct.GetProductFromSaved(productInsert);

                            // Если есть данные и нет позиции
                            if (product != null)
                            {
                                var mark = product.GetNameByRule();

                                // Если в имени нет ShortName, то добавим
                                if (!mark.Contains(product.BaseDocument.ShortName))
                                    mark = product.BaseDocument.ShortName + " " + mark;

                                // Если есть длина
                                if (product.Length != null)
                                    mark += " L=" + product.Length;
                                AddLeaderMark(ent, mark);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionBox.Show(ex);
                    }
                    finally
                    {
                        tr.Commit();
                    }
                }
            }
        }

        private static void AddLeaderMark(Entity ent, string posTxt)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            // Получаем координаты объекта
            var ext = ent.GeometricExtents;

            // Получаем координату нужной точки
            var pt = new Point3d(
                (ext.MinPoint.X + ext.MaxPoint.X) / 2,
                (ext.MinPoint.Y + ext.MaxPoint.Y) / 2,
                (ext.MinPoint.Z + ext.MaxPoint.Z) / 2);

            // Создаем текст
            var jig = new MLeaderJig
            {
                FirstPoint = pt,
                MlText = posTxt
            };
            var res = ed.Drag(jig);
            if (res.Status == PromptStatus.OK)
            {
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    btr.AppendEntity(jig.MLeader());
                    tr.AddNewlyCreatedDBObject(jig.MLeader(), true);
                    tr.TransactionManager.QueueForGraphicsFlush();
                    tr.Commit();
                }
            }
        }
    }
}