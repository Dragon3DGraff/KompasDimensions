using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using reference = System.Int32;

using KAPITypes;
using Kompas6Constants;
using Kompas6API5;
using KompasAPI7;
using Kompas6API2D5COM;


namespace KompasDimensions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        KompasObject kompas5;
        KompasObject kompas7;
        IApplication appl7;
        ksDocument2D doc5;
        IKompasDocument2D doc7;
        double[] allX = new double[0];
        double[] allY = new double[0];
        Dictionary<int, Cuts> cuts = new Dictionary<int, Cuts>();
        Dictionary<int, ContourLines> contourLines = new Dictionary<int, ContourLines>();
         Dictionary<int, CirclesInDetail> circlesInDetail = new Dictionary<int, CirclesInDetail>();
        int circleCount = 0;
        int cutsCount = 0;
        reference _contour;
        private void BtnDimensions_Click(object sender, EventArgs e)
        {
           
            string progId = string.Empty;

#if __LIGHT_VERSION__
					progId = "KOMPASLT.Application.5";
#else
            progId = "KOMPAS.Application.5";
#endif
           // try 
           // {
                kompas5 = (KompasObject)Marshal.GetActiveObject(progId);
               // MessageBox.Show(kompas5.ToString());
                if (kompas5 != null)
                {
                    kompas5.Visible = true;
                    kompas5.ActivateControllerAPI();
                }
                else
                {
                    MessageBox.Show(this, "Не найден активный объект", "Сообщение");
                }

               kompas7 = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                appl7 = (IApplication)kompas7.ksGetApplication7();
              //  appl7.Application.HideMessage = Kompas6Constants.ksHideMessageEnum.ksHideMessageYes;
               
                doc5 = (ksDocument2D)kompas5.ActiveDocument2D();
                doc7 = (IKompasDocument2D)appl7.ActiveDocument;
            double scaleOfView = 1;
                // Определим масштаб вида__________________________________
                IViews views;
            reference VN0;
            reference VNt;
            reference VN;
                IView selectedView;
                ViewsAndLayersManager viewsMng = doc7.ViewsAndLayersManager;
                views = viewsMng.Views;
                selectedView = views.ActiveView;
                
                VN0 = doc5.ksGetViewReference(selectedView.Number);
                

            int FirstViewNumber = -1;
            reference obj1;
            ksIterator iter1 = (ksIterator)kompas5.GetIterator();
            
            if (iter1.ksCreateIterator(ldefin2d.SELECT_GROUP_OBJ, VN0)) //LAYER_OBJ  //SELECT_GROUP_OBJ
            {
                if (doc5.ksExistObj(obj1 = iter1.ksMoveIterator("F")) == 1) //F - первый объект // 1 - объект существует
                {
                    do
                    {
                         FirstViewNumber = doc5.ksGetViewNumber(obj1);
                        break;
                    }
                    while (doc5.ksExistObj(obj1 = iter1.ksMoveIterator("N")) == 1); //N - следующий объект
                }
                iter1.ksDeleteIterator();
            }           
            VNt = doc5.ksGetViewReference(FirstViewNumber);
            ksViewParam Vpar = (ksViewParam)kompas5.GetParamStruct((short)StructType2DEnum.ko_ViewParam);
            doc5.ksGetObjParam(VNt, Vpar, ldefin2d.ALLPARAM);
            
            //в текущем документе и виде создадим итератор для хождения по выделенным элементам
            // IterateObj();


            //скопируем выделенное в буфер обмена

            //
            doc5.ksWriteGroupToClip(0, true); //Скопировали  выделенное в буфер 
                                                 //Создадим временный вид
                ksViewParam VparV = (ksViewParam)kompas5.GetParamStruct((short)StructType2DEnum.ko_ViewParam);
                VparV.Init();
                if (VparV != null)
                {
                    int number = 0;
                VparV.Init();
                VparV.x = 50;
                VparV.y = 50;
                VparV.scale_ = Vpar.scale_;
                VparV.angle = 0;
                VparV.color = Color.FromArgb(10, 20, 10).ToArgb();
                VparV.state = ldefin2d.stCURRENT;
                VparV.name = "Temporary";
                 scaleOfView = VparV.scale_;
                reference v = doc5.ksCreateSheetView(VparV, ref number);
                    number = doc5.ksGetViewNumber(v);
                    /*
                    ksLtVariant vart = (ksLtVariant)kompas5.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
                    if (vart != null)
                    {
                        vart.Init();
                        vart.intVal = ldefin2d.stREADONLY;
                        doc5.ksSetObjParam(v, vart, ldefin2d.VIEW_LAYER_STATE);
                    }
                    */
                }
                VN = doc5.ksGetViewReference(views.ActiveView.Number);

               // doc5.ksGetObjParam(VN, Vpar, ldefin2d.ALLPARAM);
               // Vpar.state = 3; //сделали этот вид активным
                //doc5.ksSetObjParam(VN, Vpar);
                

                //----------------------------ПОМЕНЯТЬ!!! вставить в вид ДЕТАЛИ. если вида нет - создать
               // reference pGrp = doc5.ksNewGroup(0);
                 doc5.ksStoreTmpGroup(doc5.ksReadGroupFromClip()); //Вставили все из буфера в вид детали

                                                                  
                DestroyAllObjInView(VN);
                IterateObj(VN);
                GetContour();
           // return;


            Array.Sort(allX);
            Array.Sort(allY);
                double maxX = allX[allX.Length - 1];
                double minX = allX[0];
                double maxY = allY[allY.Length - 1];
                double minY = allY[0];
                double dx = 6.5;
                double dy = 6.5;
                double dxB = 6.5;
                double dyB = 6.5;
                double dxCircle = 6.5;
                double dyCircle = 6.5;
                double dxCuts = 6.5;
                double dyCuts = 6.5;
                int direction = 1;
                short basepoint = 1;
                double displasement = 0;
            reference gr1 = doc5.ksNewGroup(0); //создали пустую группу, в неё будем складывать все полученные размеры
            doc5.ksEndGroup();


            double AxisXMiddle = minY + (maxY - minY) / 2; //средние оси. Может пригодятся
               // doc5.ksLineSeg(minX - 50, AxisXMiddle, maxX + 50, AxisXMiddle, 3);
            double AxisYMiddle = minX + (maxX - minX) / 2;
                //doc5.ksLineSeg(AxisYMiddle, minY -50, AxisYMiddle, maxY +50 , 3);
               


                if (circleCount > 0) // ПЕРЕДЕЛАТЬ
                {
                    dxB = dxB + dx;
                    dyB = dyB + dy;
                    dxCuts = dxCuts + dx;
                    dyCuts = dyCuts + dy;
                }
                if (cutsCount > 0)
                {
                    dxB = dxB + dx;
                    dyB = dyB + dy;
                }
                double minXR = Math.Round(minX, 1);
                double minYR = Math.Round(minY, 1);
                double maxXR = Math.Round(maxX, 1);
                double maxYR = Math.Round(maxY, 1);

            //********************** Образмериваем отверстия *********************************
            if (circleCount > 0)
            {
                //var YCircleMin = circlesInDetail.OrderBy(k => k.Value.YC).FirstOrDefault();
                //doc5.ksAddObjGroup(gr1, MakeDimension(minX, minY, YCircleMin.Value.XC, YCircleMin.Value.YC, (-1) * dx, 0, 1));
               // var XCircleMin = circlesInDetail.OrderBy(k => k.Value.XC).FirstOrDefault();
               // doc5.ksAddObjGroup(gr1, MakeDimension(minX, minY, XCircleMin.Value.XC, XCircleMin.Value.YC, 0, (-1) * dy, 0));
               // var ContCircle = circlesInDetail.OrderBy(k => k.Value.R).ElementAt(0);

                               //Построим горизонтальные
                double xStart = minX;
                double yStart = minY;
                double dimDirection = (-1) * dx - (yStart - minY) * Vpar.scale_;
                double rStart = circlesInDetail.OrderBy(k => k.Value.XC).ElementAt(0).Value.R;
                if(circlesInDetail.OrderBy(k => k.Value.XC).ElementAt(0).Value.YC > AxisXMiddle)
                {
                    yStart = maxY;
                    dimDirection = dx + (maxY - yStart) * Vpar.scale_;
                }
                
                //double xCurrent;
                //double yCurrent;
                /*
                int countDiameters = 1;
                double rCurrent = 0;
                for (int ir = 1; ir < circlesInDetail.Count(); ir++)
                {
                    //CirclesInDetail theCircles = circlesInDetail.OrderBy(k => k.Value.R).ElementAt(i).Value;
                   if(circlesInDetail.OrderBy(k => k.Value.R).ElementAt(ir).Value.R != rStart)
                    {
                        countDiameters++;
                    }
                        
                }
                */
               
                    for (int i = 0; i < circlesInDetail.Count(); i++)
                {                              
                    CirclesInDetail theCircles = circlesInDetail.OrderBy(k => k.Value.XC).ElementAt(i).Value;
                          if (circlesInDetail.OrderBy(k => k.Value.XC).ElementAt(i).Value.R  == rStart)
                          {
                              if (xStart != theCircles.XC)
                              {
                                doc5.ksAddObjGroup(gr1, MakeDimension(xStart, yStart, theCircles.XC, theCircles.YC, 0, dimDirection , 0));
                                xStart = theCircles.XC;
                                yStart = theCircles.YC;
                              }                             
                          }
                          else
                    {
                       
                    }
                    if (i == circlesInDetail.Count() - 1)
                    {
                        doc5.ksAddObjGroup(gr1, MakeDimension(xStart, yStart, maxX, minYR, 0, dimDirection * dx - (yStart - minY) * Vpar.scale_, 0));
                        /*
                         i = 0;
                         rStart = rCurrent;
                         xStart = minX;
                         yStart = maxY;
                         dimDirection = 1;
                         */
                    }
                }
          /*
                foreach (KeyValuePair<int, CirclesInDetail> cid in circlesInDetail)
                {
                    CirclesInDetail theCircleElement = cid.Value;
                    if (theCircleElement.Number != YCircleMin.Key)
                    {
                        //MessageBox.Show(Convert.ToString((-1) * dy - (YCircleMin.Value.YC - minY) * Vpar.scale_));
                       // doc5.ksAddObjGroup(gr1, MakeDimension(YCircleMin.Value.XC, YCircleMin.Value.YC, theCircleElement.XC, theCircleElement.YC, 0, (-1) * dx - (YCircleMin.Value.YC - minY) * Vpar.scale_, 0)); //горизонтально
                       // doc5.ksAddObjGroup(gr1, MakeDimension(YCircleMin.Value.XC, YCircleMin.Value.YC, theCircleElement.XC, theCircleElement.YC, (-1) * dy - (YCircleMin.Value.XC - minX) * Vpar.scale_, 0, 1)); //вертикально
                        //-6-70*scaleOfView
                    }

                }
                */
                circlesInDetail.Clear();
            }
            ///////////////////Образмеривание срезов///////////////////
            foreach (KeyValuePair<int, Cuts> kvp in cuts)
                {
                    Cuts theElement = kvp.Value;
                    double theElementX1R = Math.Round(theElement.X1, 1);
                    double theElementX2R = Math.Round(theElement.X2, 1);
                    double theElementY1R = Math.Round(theElement.Y1, 1);
                    double theElementY2R = Math.Round(theElement.Y2, 1);

                //для вертикальных размеров.
                
                //doc5.ksPoint(theElement.X1 + 0.5 * (theElement.X2 - theElement.X1) + 10, theElement.Y1 + 0.5*(theElement.Y2 - theElement.Y1), 1);

                if (doc5.ksIsPointInsideContour(_contour,
                        theElement.X1 + 0.5 * (theElement.X2 - theElement.X1) + 10,
                        theElement.Y1 + 0.5 * (theElement.Y2 - theElement.Y1), 1) != 0)
                {
                    if (doc5.ksIsPointInsideContour(_contour,
                        theElement.X1 + 0.5 * (theElement.X2 - theElement.X1) + 10,
                        theElement.Y1 + 0.5 * (theElement.Y2 - theElement.Y1), 1) != 1)
                    {
                        direction = -1; //строим влево
                        displasement = (theElement.X1 - minX) * scaleOfView;
                    }
                    else
                    {
                        direction = 1;
                        displasement = (theElement.X1 - maxX) * scaleOfView;
                    }

                }
                double l1 = Math.Round(Math.Abs(theElement.Y2 - theElement.Y1),3);
                double l2 = Math.Round(Math.Abs(maxY - minY),3);
                if (l1 != l2)
                { 
                    //вертикальный
                    doc5.ksAddObjGroup(gr1, MakeDimension(theElement.X1, theElement.Y1, theElement.X2, theElement.Y2, direction * dxCuts - displasement, 0, 1));
                }
                //для горизонтальных размеров.
                
                //doc5.ksCircle(theElement.X1 + 0.5 * (theElement.X2 - theElement.X1), theElement.Y1 + 0.5 * (theElement.Y2 - theElement.Y1) + 10,2, 1);
                if (doc5.ksIsPointInsideContour(_contour,
                    theElement.X1 + 0.5 * (theElement.X2 - theElement.X1),
                     theElement.Y1 + 0.5 * (theElement.Y2 - theElement.Y1) + 10, 1) != 0)
                {
                    if (doc5.ksIsPointInsideContour(_contour,
                        theElement.X1 + 0.5 * (theElement.X2 - theElement.X1),
                        theElement.Y1 + 0.5 * (theElement.Y2 - theElement.Y1) + 10, 1) != 1) //Если сверху от первой точки контур 
                    {
                        direction = -1; //строим вниз
                        displasement = (theElement.Y1 - minY) * scaleOfView;
                    }
                    else
                    {
                        direction = 1;
                        displasement = (theElement.Y1 - maxY) * Vpar.scale_;
                    }
                }
                 l1 = Math.Round(Math.Abs(theElement.X2 - theElement.X1), 3);
                 l2 = Math.Round(Math.Abs(maxX - minX), 3);
                if (l1 != l2)
                {
                    // горизонтальный
                    doc5.ksAddObjGroup(gr1, MakeDimension(theElement.X1, theElement.Y1, theElement.X2, theElement.Y2, 0, direction * dxCuts - displasement, 0, basepoint));
                     
                }                          
            }
                cuts.Clear();
           
            //*************************** Габаритные размеры *******************************
            //  MessageBox.Show(doc5.ksIsPointInsideContour(_contour, maxX, minY, 1).ToString());

            doc5.ksAddObjGroup(gr1, MakeDimension(minX, minY, maxX, minY, 0, (-1)*dxB, 0));
            doc5.ksAddObjGroup(gr1, MakeDimension(minX, minY, minX, maxY, (-1) * dyB, 0, 1));
                //kompas5.ksMessage(allY[1].ToString() + "   maxY = " + maxY.ToString() + " , minY = " + minY.ToString());
                //  kompas5.ksMessage(allX[1].ToString()+ "   maxX = " + maxX.ToString() + " , minX = " + minX.ToString()); // allX.Length.ToString());//"maxX = " + maxX.ToString() + " , minX = " + minX.ToString());
                circleCount = 0;
                cutsCount = 0;
                Array.Resize(ref allX, 0);
                Array.Resize(ref allY, 0);
            ///Array.Clear(allX, 0, 0);
            // kompas5.ksMessage(allX.Length.ToString());
            doc5.ksWriteGroupToClip(gr1, true);
           //if (doc5.ksOpenView(FirstViewNumber) == 1)
            //{
                doc5.ksOpenView(FirstViewNumber);
                doc5.ksStoreTmpGroup(doc5.ksReadGroupFromClip());
            //Удалим наш временный вид после преобразований
            doc5.ksDeleteObj(VN);
            //} 
            
            kompas5.ksMessage("Всё");
                //appl7.Application.HideMessage = kompas56Constants.ksHideMessageEnum.ksShowMessage;
           // }
            //catch
            //{
              //  MessageBox.Show("Нужно сначала запустить Компас"); // Это не нужно будет в библиотеке
            //}
        }

        private reference MakeDimension (double x1, double y1, // Привязка первой точки
                                    double x2, double y2, // привязка второй точки
                                    double dx,             // отступ от детали по x
                                    double dy,             // отступ от детали по y
                                    short gorisont,
                                    short basePoint=1)     
                                                        // размер горизонтально (gorisont = 0)
                                                        // gorisont = 1 - вертикально
                                                        // gorisont = 2 - параллельно отрезку
                                                        // gorisont = 3 - по dx, dy,
                                                        // gorisont = 4 - параллельно отрезку с выносными линиями по dx, dy.
        {             
            ksLDimParam param = (ksLDimParam)kompas5.GetParamStruct((short)StructType2DEnum.ko_LDimParam);
            if (param == null)
                return 0;

            ksDimDrawingParam dPar = (ksDimDrawingParam)param.GetDPar();
            ksLDimSourceParam sPar = (ksLDimSourceParam)param.GetSPar();
            ksDimTextParam tPar = (ksDimTextParam)param.GetTPar();

            if (dPar == null || sPar == null || tPar == null)
                return 0;
            // +++++++++++++Параметры отрисовки размера++++++++++
            dPar.Init();
            // dPar.textPos = 0; // автоматическое размещение текста (если задан textBase 1 или 2)
            dPar.textBase = 0; // текст по центру
                               // Засечки______
            dPar.pt1 = 3;
            dPar.pt2 = 3;
            // _____________
            // dPar.ang = 45; //наклон выносной линии
            //dPar.lenght = 20; //ножка выносной линии


            sPar.Init();
            // точки привязки_________
            sPar.x1 = x1;
            sPar.y1 = y1;
            sPar.x2 = x2;
            sPar.y2 = y2;
            //________________________
            // Отступ размера от детали_____________
            sPar.dx = dx; 
            sPar.dy = dy; 
            //_____________________________
            sPar.basePoint = basePoint; // от какой точки отодвигать размер 
            sPar.ps = gorisont; // размер горизонтально (0)
                         // 1 - вертикально
                         // 2 - параллельно отрезку
                         // 3 - по dx, dy,
                         // 4 - параллельно отрезку с выносными линиями по dx, dy.

            tPar.Init(false);
            tPar.SetBitFlagValue(ldefin2d._AUTONOMINAL, true);

            ksChar255 str = (ksChar255)kompas5.GetParamStruct((short)StructType2DEnum.ko_Char255);
            ksDynamicArray arrText = (ksDynamicArray)tPar.GetTextArr();

            if (str == null || arrText == null)
                return 0;
            int obj = doc5.ksLinDimension(param);
            return obj;
        }

        private void IterateObj (reference VN)
        {
            reference obj;
           
            int count = 0;
            int countContLines = 0;
            int countY = 0;
            string buf = string.Empty;
            ksIterator iter = (ksIterator)kompas5.GetIterator();
            if (iter == null)
                return;
            
                if (iter.ksCreateIterator(ldefin2d.ALL_OBJ, VN)) //LAYER_OBJ  //SELECT_GROUP_OBJ
            {
                if (doc5.ksExistObj(obj = iter.ksMoveIterator("F")) == 1) //F - первый объект // 1 - объект существует
                {
                    do
                    {
                        //   doc5.ksLightObj(obj, 1);
                       // doc5.ksLightObj(obj, 1);
                        int TypeOfObject = doc5.ksGetObjParam(obj, 0, 0); //определим тип объекта                                                                         
                        if (TypeOfObject == 1) // Если линия
                        {
                            ksLineSegParam par1 = (ksLineSegParam)kompas5.GetParamStruct((short)StructType2DEnum.ko_LineSegParam);
                            if (par1 != null)
                            {
                                int t = doc5.ksGetObjParam(obj, par1, ldefin2d.ALLPARAM);
                                if (par1.style == 1) // если она основная
                                {                                        
                                    //------- Вcе X --------
                                    if (!Array.Exists(allX, element => element == par1.x1))
                                    {
                                        Array.Resize(ref allX, allX.Length + 1);
                                        allX[count] = par1.x1;
                                        count++;
                                    }
                                    if (!Array.Exists(allX, element => element == par1.x2))
                                    {
                                        Array.Resize(ref allX, allX.Length + 1);
                                        allX[count] = par1.x2;
                                       
                                        count++;
                                    }

                                    //------- Вcе Y --------
                                    if (!Array.Exists(allY, element => element == par1.y1))
                                    {
                                        Array.Resize(ref allY, allY.Length + 1);
                                        allY[countY] = par1.y1;
                                        countY++;
                                    }
                                    if (!Array.Exists(allY, element => element == par1.y2))
                                    {
                                        Array.Resize(ref allY, allY.Length + 1);
                                        allY[countY] = par1.y2;
                                        countY++;
                                    }
                                
                                double angle = Math.Atan((par1.y2 - par1.y1) / (par1.x2 - par1.x1)) / (Math.PI / 180);

                                   // kompas5.ksMessage(Math.Round(angle).ToString());
                                    angle = Math.Round(angle);
                                    if (Math.Abs(angle) != 90.0 && angle != 0 && angle != 180 && angle != 270)
                                    {
                                    cutsCount++;
                                            Cuts theCut = new Cuts
                                            {
                                                Number = cutsCount,
                                                X1 = par1.x1,
                                                Y1 = par1.y1,
                                                X2 = par1.x2,
                                                Y2 = par1.y2
                                            };
                                            cuts.Add(key: theCut.Number, value: theCut);
                                    }
                                        ContourLines thecontourLines = new ContourLines
                                        {
                                            Number = countContLines,
                                            X1 = par1.x1,
                                            Y1 = par1.y1,
                                            X2 = par1.x2,
                                            Y2 = par1.y2
                                        };
                                        contourLines.Add(key: thecontourLines.Number, value: thecontourLines);
                                       
                                        countContLines++;
                                    }                                   
                                }
                               
                        }
                        if (TypeOfObject == 2) //если кружочек
                        {
                            ksCircleParam par1 = (ksCircleParam)kompas5.GetParamStruct((short)StructType2DEnum.ko_CircleParam);
                            if (par1 != null)
                            {
                               
                                    int t = doc5.ksGetObjParam(obj, par1, ldefin2d.ALLPARAM);
                                if (par1.style == 1) // если она основная
                                {
                                    circleCount++;
                                    CirclesInDetail thecirclesInDetail = new CirclesInDetail
                                    {
                                        Number = circleCount,
                                        R = par1.rad,
                                        XC = par1.xc,
                                        YC = par1.yc
                                    };
                                    circlesInDetail.Add(key: thecirclesInDetail.Number, value: thecirclesInDetail);
                                }

                            }
                        }
                      //  doc5.ksLightObj(obj, 0);
                    }
                    while (doc5.ksExistObj(obj = iter.ksMoveIterator("N")) == 1); //N - следующий объект
                }
                iter.ksDeleteIterator();
            } 
        }
        private void GetContour()
            {
            if (contourLines.Count != 0)
            {
            var ContL = contourLines.ElementAt(0);
            //int ContLen = contourLines.Count();
            
            double begX1 = Math.Round(ContL.Value.X1,2);
            double begY1 = Math.Round(ContL.Value.Y1,2);
            double EndX2 = Math.Round(ContL.Value.X2,2);
            double EndY2 = Math.Round(ContL.Value.Y2,2);
           // kompas5.ksMessage(begX1.ToString());
            double startX = begX1;
            double startY = begY1;
            double FinX = EndX2;
            double FinY = EndY2;
            int coincidence = 0;
            int iterations = 0;
            int iterations2 = 0;
            int i = 0;
            contourLines.Remove(ContL.Key);
           
            //contourLines.Remove(0);
           // MessageBox.Show(ContLen.ToString());
            //Построим контур
           // kompas5.ksMessage("i = " + i + ", contourLines.Count() = " + contourLines.Count() + ",  = ");
            if (doc5.ksContour(6) == 1)// 6 - это стиль линии (вспомогательная), 1 - успешное создание контура
            {
                //kompas5.ksMessage("i = " + i + ", begX1 = " + begX1 + ", begY1 = " + begY1 + ", EndX2 = " + EndX2 + ", EndY2 = " + EndY2);
                doc5.ksLineSeg(begX1, begY1, EndX2, EndY2, 6);
         
                while (FinX != startX || FinY != startY) 
                {
                    //kompas5.ksMessage(i + "  ,    startX = " + startX + " <=> " + FinX + "= FinX " + '\n' + "startY = " + startY + " <=> " + FinY + " = FinY");

                    System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
                    foreach (KeyValuePair<int, ContourLines> cid in contourLines)
                    {
                        ContourLines thecontourLines = cid.Value;

                        messageBoxCS.AppendFormat("{0} = {1}", "Number", thecontourLines.Number);
                        messageBoxCS.AppendLine();
                        messageBoxCS.AppendFormat("{0} = {1}", "X1", thecontourLines.X1);
                        messageBoxCS.AppendLine();
                        messageBoxCS.AppendFormat("{0} = {1}", "Y1", thecontourLines.Y1);
                        messageBoxCS.AppendLine();
                        messageBoxCS.AppendFormat("{0} = {1}", "X2", thecontourLines.X2);
                        messageBoxCS.AppendLine();
                        messageBoxCS.AppendFormat("{0} = {1}", "Y2", thecontourLines.Y2);
                        messageBoxCS.AppendLine();
                    }
                    //kompas5.ksMessage(messageBoxCS.ToString());


                    ContL = contourLines.ElementAt(i);
                   //  kompas5.ksMessage("i = " + i + ", EndX2 = " + EndX2 + ", ContL.Value.X1 = " + ContL.Value.X1);
                    //double chX1 = Math.Round(ContL.Value.X1, 2);
                    //double chX2 = Math.Round(ContL.Value.X2, 2);
                    //kompas5.ksMessage("i = " + i + ", EndX2 = " + EndX2 + ", chX1 = " + chX1  + ", chX2 = " + chX2);
                    if (EndX2 == Math.Round(ContL.Value.X1, 2) && EndY2 == Math.Round(ContL.Value.Y1, 2))
                    {
                        // kompas5.ksMessage("Я тут!");

                        // kompas5.ksMessage("В Y1 зашел!");
                        begX1 = Math.Round(ContL.Value.X1, 2);
                        begY1 = Math.Round(ContL.Value.Y1, 2);
                        EndX2 = Math.Round(ContL.Value.X2, 2);
                        EndY2 = Math.Round(ContL.Value.Y2, 2);
                        FinX = EndX2;
                        FinY = EndY2;
                      //  kompas5.ksMessage("i = " + i + "В Y1 зашел! begX1 = " + begX1 + ", begY1 = " + begY1 + ", EndX2 = " + EndX2 + ", EndY2 = " + EndY2);
                        doc5.ksLineSeg(begX1, begY1, EndX2, EndY2, 6);
                        coincidence++;
                        contourLines.Remove(ContL.Key);
                        // kompas5.ksMessage("i = " + i + ", contourLines.Count() = " + contourLines.Count() + ",  = ");

                    }
                    else
                    {
                        if (EndX2 == Math.Round(ContL.Value.X2, 2) && EndY2 == Math.Round(ContL.Value.Y2, 2))
                        {
                            // kompas5.ksMessage("Или тут!");

                            // kompas5.ksMessage("В Y2 зашел!");

                            begX1 = Math.Round(ContL.Value.X2, 2);
                            begY1 = Math.Round(ContL.Value.Y2, 2);
                            EndX2 = Math.Round(ContL.Value.X1, 2);
                            EndY2 = Math.Round(ContL.Value.Y1, 2);
                            FinX = EndX2;
                            FinY = EndY2;
                          //   kompas5.ksMessage("i = " + i + ", В Y2 зашел! begX1 = " + begX1 + ", begY1 = " + begY1 + ", EndX2 = " + EndX2 + ", EndY2 = " + EndY2);
                            doc5.ksLineSeg(begX1, begY1, EndX2, EndY2, 6);
                            coincidence++;
                            contourLines.Remove(ContL.Key);
                            // kompas5.ksMessage("i = " + i + ", contourLines.Count() = " + contourLines.Count() + ",  = ");

                        }
                    }
                     //kompas5.ksMessage(i + "  contourLines.Count() = " + contourLines.Count() + ", ContL.Value.X2 = " + ContL.Value.X2 + ", ContL.Value.Y2 = " + ContL.Value.Y2 + ",    ContL.Value.X1 = " + ContL.Value.X1 + ", ContL.Value.Y1 = " + ContL.Value.Y1);

                   // kompas5.ksMessage(i + "  contourLines.Count() = " + contourLines.Count() + ", FinX = " + FinX + ", FinY = " + FinY + ",    startX = " + startX + ", startY = " + startY);
                    if (i >= contourLines.Count() - 1)
                    {
                        i = 0;
                        if (contourLines.Count() == 0) break;
                        if (coincidence == 0)
                        {
                            iterations++;
                            iterations2++;
                            //kompas5.ksMessage("Нет совпадений, iterations = " + iterations);
                        }
                        else
                        {
                            iterations = 0;
                            coincidence = 0;
                        }
                        if (iterations == 4)
                        {
                            kompas5.ksMessage("Контур не замкнут! Проверь размеры!!!");
                            double TempStartX = startX;
                            double TempStartY = startY;
                            startX = FinX;
                            startY = FinY;
                            FinX = TempStartX;
                            FinY = TempStartY;
                            EndX2 = FinX;
                            EndY2 = FinY;
                            
                            //coincidence = 0;
                        }
                        else if (iterations2 > 8)
                        {
                           
                            break;
                        }
                        // i++;
                    }
                    else
                    {
                        i++;
                       // coincidence = 0;
                    }
                    if (FinX != startX && FinY != startY)
                    {

                    }
                }
                
                doc5.ksLineSeg(FinX, FinY, startX, startY, 6);
                // doc5.ksLineSeg(par1.x1, par1.y1, par1.x2, par1.y2, 6); // это для контура
                // doc5.ksLightObj(obj, 1);

                // doc5.ksLightObj(obj, 0);
                _contour = doc5.ksEndObj();
                contourLines.Clear();
              //  doc5.ksLightObj(_contour, 1);
              //  kompas5.ksMessage("");
              //  doc5.ksLightObj(_contour, 0);
            }
            }
        }
        private void DestroyAllObjInView(reference VN)
        {
            reference obj;
            ksIterator iter = (ksIterator)kompas5.GetIterator();
            if (iter == null)
                return;
            if (iter.ksCreateIterator(ldefin2d.ALL_OBJ, VN))
            {
                //создадим итератор для хождения по виду
                obj = iter.ksMoveIterator("F");
                if (doc5.ksExistObj(obj) == 1)
                {
                    do
                    {
                        int TypeOfObject = doc5.ksGetObjParam(obj, 0, 0); //определим тип объекта
                        if (TypeOfObject == 26 || //контур
                            TypeOfObject == 27 || //нетипи­зиро­ванный макро­эле­мент                          
                            TypeOfObject == 31 || //ломаная линия
                            TypeOfObject == 35 || //прямоугольник
                            TypeOfObject == 36 || //правильный многоугольник
                            TypeOfObject == 72)   //мультилиния
                        {
                            // doc5.ksLightObj(obj, 1);
                            if (TypeOfObject == 27)
                            {
                               // doc5.ksLightObj(obj, 1);
                                //double xM =0;
                              //  double yM = 0;
                              //  double AngM = 0;
                                //long Y = doc5.ksGetMacroParam(obj,0);
                                //kompas5.ksMessage(Y.ToString());
                                // doc5 = (ksDocument2D)kompas5.Document2D();

                                // doc5.ksGetMacroParamSize(obj);
                                //IDrawingContainer drawCont = GetDrawingContainer();

                                // IDrawingObject baseObj = (IDrawingObject)kompas.TransferInterface(obj, doc5.reference);
                                // UserParam userParam = (UserParam)kompas5.GetParamStruct((short)StructType2DEnum.ko_UserParam);
                                //int op = doc5.ksGetMacroParam(obj, userParam);

                                //kompas5.ksMessage(op.ToString());
                               // IDrawingObject drawingObject; //= (IDrawingObject)kompas5.TransferReference(obj, 0);// doc5.reference);
                              

                                //kompas7.tr
                                // doc5.ma
                               // doc5.ksGetMacroPlacement(obj, ref  xM, ref  yM, ref  AngM);
                                //kompas5.ksMessage("xM = " + xM.ToString() + ", yM = " + yM + ", angM = " + AngM);
                               // doc5.ksLightObj(obj, 0);
                            }
                            int result = doc5.ksDestroyObjects(obj); //Разрушим  
                           // kompas5.ksMessage("result = " + result);
                            DestroyAllObjInView(VN);
                        }

                    }
                    while (doc5.ksExistObj(obj = iter.ksMoveIterator("N")) == 1);
                }
                iter.ksDeleteIterator();
            }

        }
        //***********************************************ПРОСТАНОВКА СВАРКИ*********************************************

        private void button1_Click(object sender, EventArgs e)
        {
            string progId = string.Empty;

#if __LIGHT_VERSION__
					progId = "KOMPASLT.Application.5";
#else
            progId = "KOMPAS.Application.5";
#endif
           
                kompas5 = (KompasObject)Marshal.GetActiveObject(progId);
                // MessageBox.Show(kompas5.ToString());
                if (kompas5 != null)
                {
                    kompas5.Visible = true;
                    kompas5.ActivateControllerAPI();
                }
                else
                {
                    MessageBox.Show(this, "Не найден активный объект", "Сообщение");
                }

                kompas7 = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                appl7 = (IApplication)kompas7.ksGetApplication7();
                //  appl7.Application.HideMessage = Kompas6Constants.ksHideMessageEnum.ksHideMessageYes;

                doc5 = (ksDocument2D)kompas5.ActiveDocument2D();
                doc7 = (IKompasDocument2D)appl7.ActiveDocument;

            // Определим масштаб вида__________________________________
            IViews views;
            reference VN;
            IView selectedView;
            ViewsAndLayersManager viewsMng = doc7.ViewsAndLayersManager;
            views = viewsMng.Views;
            selectedView = views.ActiveView;
            ksViewParam Vpar = (ksViewParam)kompas5.GetParamStruct((short)StructType2DEnum.ko_ViewParam);
            VN = doc5.ksGetViewReference(selectedView.Number);
            doc5.ksGetObjParam(VN, Vpar, ldefin2d.ALLPARAM);
            double scaleOfView = Vpar.scale_;

            reference obj;
            ksIterator iter = (ksIterator)kompas5.GetIterator();
            if (iter == null)
                return;
            if (iter.ksCreateIterator(ldefin2d.SELECT_GROUP_OBJ, VN)) //LAYER_OBJ  //SELECT_GROUP_OBJ
            {
                if (doc5.ksExistObj(obj = iter.ksMoveIterator("F")) == 1) //F - первый объект // 1 - объект существует
                {
                    do
                    {
                        //   doc5.ksLightObj(obj, 1);

                        int TypeOfObject = doc5.ksGetObjParam(obj, 0, 0); //определим тип объекта
                                                                          // MessageBox.Show(TypeOfObject.ToString());
                                                                          // kompas5.ksMessage(TypeOfObject.ToString());

                        if (TypeOfObject == 1) //
                        {
                            ksLineSegParam par1 = (ksLineSegParam)kompas5.GetParamStruct((short)StructType2DEnum.ko_LineSegParam);
                            if (par1 != null)
                            {
                                int t = doc5.ksGetObjParam(obj, par1, ldefin2d.ALLPARAM);
                                //Угол наклона линии
                                double angle = Math.Atan((par1.y2 - par1.y1) / (par1.x2 - par1.x1)) / (Math.PI / 180);
                                //kompas5.ksMessage(angle.ToString());
                                // введем локальную систему координат (матрица преобразований это)
                                doc5.ksMtr(par1.x1+0.5 *(par1.x2- par1.x1), par1.y1 + 0.5 * (par1.y2 - par1.y1), angle, 1, 1);

                                //это значок сварки макроэлементом
                                doc5.ksMacro(0);
                                doc5.ksLineSeg(0, 0, 0, 2 / scaleOfView, 1);                              
                                doc5.ksArcByAngle(0, 0, 1.25 / scaleOfView, 0, 90, 1, 1);
                                doc5.ksEndObj();

                                //Удалим матрицу преобразований, иначе они складываются и получается очень весело
                                  doc5.ksDeleteMtr();
                            }
                        }
                    }
                    while (doc5.ksExistObj(obj = iter.ksMoveIterator("N")) == 1); //N - следующий объект
                    
                    }
            
                iter.ksDeleteIterator();
               
            }
        }
    }
    public class Cuts
    {
        public int Number { get; set; }
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
    }
    public class CirclesInDetail
    {
        public int Number { get; set; }
        public double R { get; set; }
        public double XC { get; set; }
        public double YC { get; set; }        
    }
    public class ContourLines
    {
        public int Number { get; set; }
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
    }
}
