using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExerciseTest1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (NearPoint.Text != null)
                NearPoint.Text = "";
            double segmentLength1 = segmentLength(Convert.ToDouble(StartPointX.Text), Convert.ToDouble(FirstPointX.Text),
                                                  Convert.ToDouble(StartPointY.Text), Convert.ToDouble(FirstPointZ.Text),
                                                  Convert.ToDouble(StartPointZ.Text), Convert.ToDouble(FirstPointY.Text));
            double segmentLength2 = segmentLength(Convert.ToDouble(StartPointX.Text), Convert.ToDouble(SecondPointX.Text),
                                                  Convert.ToDouble(StartPointY.Text), Convert.ToDouble(SecondPointY.Text),
                                                  Convert.ToDouble(StartPointZ.Text), Convert.ToDouble(SecondPointZ.Text));
            double segmentLength3 = segmentLength(Convert.ToDouble(StartPointX.Text), Convert.ToDouble(ThirdPointX.Text),
                                                  Convert.ToDouble(StartPointY.Text), Convert.ToDouble(ThirdPointY.Text),
                                                  Convert.ToDouble(StartPointZ.Text), Convert.ToDouble(ThirdPointZ.Text));
            double segmentLength4 = segmentLength(Convert.ToDouble(StartPointX.Text), Convert.ToDouble(FourthPointX.Text),
                                                  Convert.ToDouble(StartPointY.Text), Convert.ToDouble(FourthPointY.Text),
                                                  Convert.ToDouble(StartPointZ.Text), Convert.ToDouble(FourthPointZ.Text));
            double[] segments = { segmentLength1, segmentLength2, segmentLength3, segmentLength4 };
            double minSegment = segments[0];
            for (int i = 0; i < segments.Length; i++)
            {
                if (segments[i] < minSegment)
                {
                    minSegment = segments[i];
                }
            }
            NearPoint.Text = Convert.ToString(minSegment);
        }
        double segmentLength(double X1, double X2,
                         double Y1, double Y2,
                         double Z1, double Z2)
        {
            double X = X1 - X2;
            double Y = Y1 - Y2;
            double Z = Z1 - Z2;
            double tempLength = (X * X) + (Y * Y) + (Z * Z);
            return Math.Sqrt(tempLength);
        }
        

        private void button2_Click(object sender, EventArgs e)
        {
            myPoint start = new myPoint(Convert.ToDouble(StartPointX.Text) / 1000, Convert.ToDouble(StartPointY.Text) / 1000, Convert.ToDouble(StartPointZ.Text) / 1000);
            myPoint final = new myPoint(Convert.ToDouble(FinalPointX.Text) / 1000, Convert.ToDouble(FinalPointY.Text) / 1000, Convert.ToDouble(FinalPointZ.Text) / 1000);
            myPoint A = new myPoint(Convert.ToDouble(FirstPointX.Text) / 1000, Convert.ToDouble(FirstPointY.Text) / 1000, Convert.ToDouble(FirstPointZ.Text) / 1000);
            myPoint B = new myPoint(Convert.ToDouble(SecondPointX.Text) / 1000, Convert.ToDouble(SecondPointY.Text) / 1000, Convert.ToDouble(SecondPointZ.Text) / 1000);
            myPoint C = new myPoint(Convert.ToDouble(ThirdPointX.Text) / 1000, Convert.ToDouble(ThirdPointY.Text) / 1000, Convert.ToDouble(ThirdPointZ.Text) / 1000);
            myPoint D = new myPoint(Convert.ToDouble(FourthPointX.Text) / 1000, Convert.ToDouble(FourthPointY.Text) / 1000, Convert.ToDouble(FourthPointZ.Text) / 1000);
            myPoint findpointA = FindProjectPoint(start, final, A);
            myPoint findpointB = FindProjectPoint(start, final, B);
            myPoint findpointC = FindProjectPoint(start, final, C);
            myPoint findpointD = FindProjectPoint(start, final, D);

            ResultPoint OA = new ResultPoint(distance(start, findpointA), "A");
            ResultPoint OB = new ResultPoint(distance(start, findpointB), "B");
            ResultPoint OC = new ResultPoint(distance(start, findpointC), "C");
            ResultPoint OD = new ResultPoint(distance(start, findpointD), "D");

            ResultPoint[] segments = { OA, OB, OC, OD };
            ResultPoint nearestPoint = segments[0];
            for (int i = 0; i < segments.Length; i++)
            {
                if (segments[i].X < nearestPoint.X)
                {
                    nearestPoint = segments[i];
                }
            }
            MessageBox.Show(Convert.ToString(nearestPoint.Name),"Ближайшая точка к началу отрезка:");
        }
        myPoint FindProjectPoint(myPoint start, myPoint final, myPoint point)
        {
            //Проверяю есть ли солид, и запускаю построения
            SldWorks SwApp;
            IModelDoc2 model;
            SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.30");
            SwApp.NewDocument(@"C:\SWPlus_v_2017_SP0.0\Шаблоны\gost-part.prtdot", 0, 0, 0);
            model = (ModelDoc2)SwApp.ActiveDoc;
            //Строю линию и точку
            model.SketchManager.Insert3DSketch(true);
            model.SketchManager.CreateLine(start.X, start.Y, start.Z, final.X, final.Y, final.Z);
            model.SketchManager.CreatePoint(point.X, point.Y, point.Z);
            model.SketchManager.Insert3DSketch(true);
            //Строю плоскости
            bool boolstatus0 = model.Extension.SelectByID2("Point1@Трехмерный эскиз1", "EXTSKETCHPOINT", start.X, start.Y, start.Z, true, 0, null, 0);
            bool boolstatus1 = model.Extension.SelectByID2("Point2@Трехмерный эскиз1", "EXTSKETCHPOINT", final.X, final.Y, final.Z, true, 1, null, 0);
            bool boolstatus2 = model.Extension.SelectByID2("Point3@Трехмерный эскиз1", "EXTSKETCHPOINT", point.X, point.Y, point.Z, true, 2, null, 0);
            model.FeatureManager.InsertRefPlane(4, 0, 4, 0, 4, 0);

            bool boolstatus3 = model.Extension.SelectByID2("Point1@Трехмерный эскиз1", "EXTSKETCHPOINT", start.X, start.Y, start.Z, true, 0, null, 0);
            bool boolstatus4 = model.Extension.SelectByID2("Point2@Трехмерный эскиз1", "EXTSKETCHPOINT", final.X, final.Y, final.Z, true, 1, null, 0);
            bool boolstatus5 = model.Extension.SelectByID2("Плоскость1", "PLANE", point.X, point.Y, point.Z, true, 2, null, 0);
            model.FeatureManager.InsertRefPlane(4, 0, 4, 0, 2, 0);
            //проецирую точку

            bool boolstatus6 = model.Extension.SelectByID2("Плоскость2", "PLANE", 0, 0, 0, false, 0, null, 0);
            model.SketchManager.InsertSketch(true);
            bool boolstatus7 = model.Extension.SelectByID2("Point3@Трехмерный эскиз1", "EXTSKETCHPOINT", point.X, point.Y, point.Z, true, 2, null, 0);
            bool boolstatus8 = model.SketchManager.SketchUseEdge3(false, false);

            model.SketchManager.InsertSketch(true);
            //Получаю координаты спроецированной точки в плоскости

            SelectionMgr swSelMgr = (SelectionMgr)model.SelectionManager;
            Feature swFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);
            Sketch swSketch = (Sketch)swFeat.GetSpecificFeature2();
            object[] sketchPoints = (object[])swSketch.GetSketchPoints2();
            SketchPoint swSketchPt = (SketchPoint)sketchPoints[0];

            SelectData swSelData = (SelectData)swSelMgr.CreateSelectData();

            bool status = swSketchPt.Select4(true, swSelData);
            //Создаю 3D эскиз и проецирую точку в 3D пространство, так пришлось делать - не нашел другой API команды

            model.SketchManager.Insert3DSketch(true);
            bool boolstatus9 = model.Extension.SelectByID2("Point1Эскиз1", "EXTSKETCHPOINT", swSketchPt.X / 1000, swSketchPt.Y / 1000, swSketchPt.Z / 1000, false, 0, null, 0);
            bool boolstatus10 = model.SketchManager.SketchUseEdge3(false, false);
            model.ClearSelection2(true);
            model.SketchManager.Insert3DSketch(true);

            SelectionMgr swSelMgr2 = (SelectionMgr)model.SelectionManager;
            Feature swFeat2 = (Feature)swSelMgr2.GetSelectedObject6(1, -1);
            Sketch swSketch2 = (Sketch)swFeat2.GetSpecificFeature2();
            object[] sketchPoints2 = (object[])swSketch2.GetSketchPoints2();
            SketchPoint swSketchPt2 = (SketchPoint)sketchPoints2[0];
            model.ClearSelection2(true);
            myPoint findPoint = new myPoint(swSketchPt2.X*1000, swSketchPt2.Y*1000, swSketchPt2.Z*1000);
            //SwApp.CloseDoc(model.GetPathName().ToString());
            return findPoint;
        }
        double distance(myPoint A, myPoint B)
        {
            double X = B.X - A.X;
            double Y = B.Y - A.Y;
            double Z = B.Z - A.Z;
            double tempLength = (X * X) + (Y * Y) + (Z * Z);
            return Math.Sqrt(tempLength);
        }
    }
}
