//*********************************************************************
//Copyright(C) 2021 EDAL solutions BV
//URL: https://www.edalsolutions.be
//Using XCAD from Xarial
//*********************************************************************

using EquiDistAddin.Properties;
using System;
using static System.Environment;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Xarial.XCad.Base.Attributes;
using Xarial.XCad.UI.Commands.Attributes;
using Xarial.XCad.UI.Commands;

using Xarial.XCad.SolidWorks;
using Xarial.XCad.UI.PropertyPage;

using Xarial.XCad.UI.PropertyPage.Enums;
using System.Drawing;

using Xarial.XCad.UI.Commands.Enums;
using Xarial.XCad.SolidWorks.UI.PropertyPage;
using System.Diagnostics;
using Xarial.XCad.UI.PropertyPage.Attributes;
using Xarial.XCad.Base.Enums;
using SolidWorks.Interop.sldworks;
using System.Windows.Media.Media3D;

namespace EquiDistAddin
{

    #region ENUMS

    //--- Define the command manager buttons
    [Title("EQUIDIST")]//this name will be shown in the CommandManager TAB
    [Description("Equidistant Points on a Hemisphere (The Thomson Problem)")]
    [CommandGroupInfo(10)]
    public enum Commands_ed
    {
        [Title("Equidistant points on sphere")]
        [Description("Creates Equidistant points on a sphere")]
        [Icon(typeof(Resources), nameof(Resources.points_simple))]
        [CommandItemInfo(true, true, WorkspaceTypes_e.Part, true)]
        CommandB1,

        [Title("Info")]
        [Description("Documentation and Credits")]
        [Icon(typeof(Resources), nameof(Resources.about))]
        CommandB2

    }

    //--- Define the sphere optionboxes
    public enum Options_sphere
    {
        [Title("Hemisphere")]
        //[BitmapOptions(18, 18)]
        [Icon(typeof(Resources), nameof(Resources.hemisphere))]
        Hemisphere,

        [Title("Sphere")]
        //[BitmapOptions(18, 18)]
        [Icon(typeof(Resources), nameof(Resources.sphere))]
        Sphere

    }

    #endregion

    //--- Define the PropertyManager Page
    #region PMP

    //define the controls you want in the PMP
    public class DataModelSphereOptions
    {

        [BitmapOptions(220, 220)]
        public Image BitmapLarge { get; set; } = Resources.sphere_illustration;

        [NumberBoxOptions(NumberBoxUnitType_e.Length, 0.1, 10.0, 0.1, true, 1, 0.1, NumberBoxStyle_e.Thumbwheel)]//in m!
        [Description("Radius of the sphere in mm")]
        [StandardControlIcon(BitmapLabelType_e.Radius)]
        public double Radius { get; set; } = 1.0;

        [NumberBoxOptions(NumberBoxUnitType_e.UnitlessInteger, 1, 6, 1, true, 2, 1, NumberBoxStyle_e.Slider)]
        [Description("Number of recursions. This is how many times we want to divide the distance of the points. The image above shows the result with 2 recursions")]
        [StandardControlIcon(BitmapLabelType_e.CircularPattern)]
        public int Recursions { get; set; } = 1;


        public class DataGroupInput
        {

            [Title("Sphere type")]
            [Description("Choose a Hemisphere or generate the complete Sphere")]
            [OptionBox]
            //put the following separately for better code readability
            //public enum Options_sphere
            //{
            //    Hemisphere,
            //    Sphere
            //}
            [ComboBoxOptions(ComboBoxStyle_e.Sorted)]
            public Options_sphere SphereOptions { get; set; } = Options_sphere.Hemisphere;

            ////TODO ADD ACTIONS
            //[BitmapButton(typeof(Resources), nameof(Resources.hemisphere))]
            //public bool ToggleHemisphere { get; set; } = true;

            //[BitmapButton(typeof(Resources), nameof(Resources.sphere))]
            //public bool ToggleSphere { get; set; } = false;

            [Title("Draw Normals")]
            [Description("Draw the normals on the sphere for every point")]
            [Icon(typeof(Resources), nameof(Resources.normals_sphere))]
            //[ControlOptions(height: 16, width: 16)]//doesn't do anything
            public bool NormalsChecked { get; set; } = false;

        }

        [Title("Options")]
        public DataGroupInput Group1 { get; set; }


        ////draw spike
    }

    //PMP handler
    [ComVisible(true)]
    [Title("Equidistant points on sphere")]
    public class MyPMPageData : SwPropertyManagerPageHandler
    {
        [Title("Dimensions")]
        public DataModelSphereOptions InputSpherePoints { get; set; }

    }

    #endregion

    //---
    #region ADDIN

    [ComVisible(true), Guid("15764d0d-35fb-419c-b1c0-54821484f5d5")]
    [Icon(typeof(Resources), nameof(Resources.swcomaddinwizard))]
    [Title("EquiDistant Points Addin")]
    [Description("SWPUC32 - Equidistant Points on a Hemisphere (The Thomson Problem)")]
    public class AddIn : SwAddInEx
    {
        //--- private members
        private IXPropertyPage<MyPMPageData> m_Page;
        private MyPMPageData m_Data;

        //--- On starting SolidWorks
        public override void OnConnect()
        {

            CommandManager.AddCommandGroup<Commands_ed>().CommandClick += OnCommandClick;

            //initialize the PMP data
            m_Data = new MyPMPageData();


        }

        //--- handle the command manager buttons
        private void OnCommandClick(Commands_ed spec)
        {
            switch (spec)
            {
                case Commands_ed.CommandB1:
                    m_Page = this.CreatePage<MyPMPageData>();
                    ShowPmpPage();
                    break;

                case Commands_ed.CommandB2:
                    string msg = "The tool creates a 3Dsketch and puts points on the surface of a hemisphere." + NewLine +
                                  NewLine +
                                 "The points are calculated so that they get an ideal equidistant gap between them." + NewLine +
                                 "The user can control the radius from 100 mm to 10 m and the recursion level from 1 to 6." + NewLine +
                                 "Each time you run the command, a new 3Dsketch is created, so you can use multiple at once." + NewLine +
                                  NewLine +
                                 "We used XCAD from https://xcad.xarial.com/ " + NewLine +
                                  NewLine +
                                 "and " + NewLine +
                                  NewLine +
                                 "tweaked code from the following website: " + NewLine +
                                 "http://blog.andreaskahler.com" + NewLine +
                                  NewLine +
                                 "Author: Eddy Alleman - https://edalsolutions.be/";

                    Application.ShowMessageBox(msg, MessageBoxIcon_e.Info, MessageBoxButtons_e.Ok);
                    // Application.ShowMessageBox("Button2");
                    break;

            }

        }

        //--- Helpers
        private void ShowPmpPage()
        {
            m_Page.Closed += OnPageClosed;
            m_Page.Show(m_Data);
        }

        //--- debugging
        private void OnPageClosed(PageCloseReasons_e reason)
        {


            if (reason == PageCloseReasons_e.Okay)
            {

                //get input from PMP
                double rad = m_Data.InputSpherePoints.Radius;
                Debug.Print($"Radius: {rad}");

                int recur = m_Data.InputSpherePoints.Recursions;
                Debug.Print($"Recursions: {recur}");

                Options_sphere hemi = m_Data.InputSpherePoints.Group1.SphereOptions;
                Debug.Print($"Type of sphere: {hemi}");

                bool normals = m_Data.InputSpherePoints.Group1.NormalsChecked;
                Debug.Print($"Normals on sphere: {normals}");


                ISldWorks swApp = this.Application.Sw;

                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;

                if (swModel is PartDoc)
                {
                    IcoSphereCreator icoSphere = new IcoSphereCreator();

                    MeshGeometry3D meshData = new MeshGeometry3D();
                    meshData = icoSphere.Create(rad, recur);//tested up to 5

                    //get vertex posiitons
                    Point3DCollection positions = meshData.Positions;

                    int nPositions = positions.Count;

                    SketchManager sketchMgr = swModel.SketchManager;
                    sketchMgr.AddToDB = true;
                    sketchMgr.Insert3DSketch(true);

                    CreatePointsAndNormals(rad, hemi, normals, positions, nPositions, sketchMgr);

                    //close 3Dsketch
                    sketchMgr.InsertSketch(true);

                    swModel.ViewZoomtofit2();
                }

            }

            #endregion

        }

        private static void CreatePointsAndNormals(double rad, Options_sphere hemi, bool normals, Point3DCollection positions, int nPositions, SketchManager sketchMgr)
        {
            //normals length factor to radius
            double factor = 1.2 * rad;

            //define limit value in Y direction. everything above tha value will be shown. 
            double ceiling = (hemi == Options_sphere.Hemisphere) ? 0.0 : -1.0;


            for (int i = 0; i < nPositions; i++)
            {

                if (positions[i].Y >= ceiling)
                {

                    if (normals)
                    {
                        sketchMgr.CreateLine(rad * positions[i].X, rad * positions[i].Y, rad * positions[i].Z, factor * positions[i].X, factor * positions[i].Y, factor * positions[i].Z);
                    }
                    else
                    {
                        SketchPoint pt = sketchMgr.CreatePoint(rad * positions[i].X, rad * positions[i].Y, rad * positions[i].Z);
                    }
                }
            }
        }

    }
}
