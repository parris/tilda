using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks
{
    class MockShapes : Shapes
    {
        private List<MockShape> shapes = new List<MockShape>();

        public Shape AddCallout(Microsoft.Office.Core.MsoCalloutType Type, float Left, float Top, float Width, float Height)
        {
            throw new NotImplementedException();
        }

        public Shape AddCanvas(float Left, float Top, float Width, float Height)
        {
            throw new NotImplementedException();
        }

        public Shape AddChart(Microsoft.Office.Core.XlChartType Type, float Left = -1f, float Top = -1f, float Width = -1f, float Height = -1f)
        {
            throw new NotImplementedException();
        }

        public Shape AddComment(float Left = 1.25f, float Top = 1.25f, float Width = 145.25f, float Height = 145.25f)
        {
            throw new NotImplementedException();
        }

        public Shape AddConnector(Microsoft.Office.Core.MsoConnectorType Type, float BeginX, float BeginY, float EndX, float EndY)
        {
            throw new NotImplementedException();
        }

        public Shape AddCurve(object SafeArrayOfPoints)
        {
            throw new NotImplementedException();
        }

        public Shape AddDiagram(Microsoft.Office.Core.MsoDiagramType Type, float Left, float Top, float Width, float Height)
        {
            throw new NotImplementedException();
        }

        public Shape AddLabel(Microsoft.Office.Core.MsoTextOrientation Orientation, float Left, float Top, float Width, float Height)
        {
            throw new NotImplementedException();
        }

        public Shape AddLine(float BeginX, float BeginY, float EndX, float EndY)
        {
            throw new NotImplementedException();
        }

        public Shape AddMediaObject(string FileName, float Left = 0f, float Top = 0f, float Width = -1f, float Height = -1f)
        {
            throw new NotImplementedException();
        }

        public Shape AddMediaObject2(string FileName, Microsoft.Office.Core.MsoTriState LinkToFile = Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState SaveWithDocument = Microsoft.Office.Core.MsoTriState.msoTrue, float Left = 0f, float Top = 0f, float Width = -1f, float Height = -1f)
        {
            throw new NotImplementedException();
        }

        public Shape AddMediaObjectFromEmbedTag(string EmbedTag, float Left = 0f, float Top = 0f, float Width = -1f, float Height = -1f)
        {
            throw new NotImplementedException();
        }

        public Shape AddOLEObject(float Left = 0f, float Top = 0f, float Width = -1f, float Height = -1f, string ClassName = "", string FileName = "", Microsoft.Office.Core.MsoTriState DisplayAsIcon = Microsoft.Office.Core.MsoTriState.msoFalse, string IconFileName = "", int IconIndex = 0, string IconLabel = "", Microsoft.Office.Core.MsoTriState Link = Microsoft.Office.Core.MsoTriState.msoFalse)
        {
            throw new NotImplementedException();
        }

        public Shape AddPicture(string FileName, Microsoft.Office.Core.MsoTriState LinkToFile, Microsoft.Office.Core.MsoTriState SaveWithDocument, float Left, float Top, float Width = -1f, float Height = -1f)
        {
            throw new NotImplementedException();
        }

        public Shape AddPlaceholder(PpPlaceholderType Type, float Left = -1f, float Top = -1f, float Width = -1f, float Height = -1f)
        {
            throw new NotImplementedException();
        }

        public Shape AddPolyline(object SafeArrayOfPoints)
        {
            throw new NotImplementedException();
        }

        public Shape AddShape(Microsoft.Office.Core.MsoAutoShapeType Type, float Left, float Top, float Width, float Height)
        {
            throw new NotImplementedException();
        }

        public Shape AddSmartArt(Microsoft.Office.Core.SmartArtLayout Layout, float Left = -1f, float Top = -1f, float Width = -1f, float Height = -1f)
        {
            throw new NotImplementedException();
        }

        public Shape AddTable(int NumRows, int NumColumns, float Left = -1f, float Top = -1f, float Width = -1f, float Height = -1f)
        {
            throw new NotImplementedException();
        }

        public Shape AddTextEffect(Microsoft.Office.Core.MsoPresetTextEffect PresetTextEffect, string Text, string FontName, float FontSize, Microsoft.Office.Core.MsoTriState FontBold, Microsoft.Office.Core.MsoTriState FontItalic, float Left, float Top)
        {
            throw new NotImplementedException();
        }

        public Shape AddTextbox(Microsoft.Office.Core.MsoTextOrientation Orientation, float Left, float Top, float Width, float Height)
        {
            MockShape shape = new MockShape();
            shape.type = Microsoft.Office.Core.MsoShapeType.msoTextBox;
            shape.Left = Left;
            shape.Top = Top;
            shape.Width = Width;
            shape.Height = Height;
            shapes.Add(shape);
            return shape;
        }

        public Shape AddTitle()
        {
            throw new NotImplementedException();
        }

        public dynamic Application
        {
            get { throw new NotImplementedException(); }
        }

        public FreeformBuilder BuildFreeform(Microsoft.Office.Core.MsoEditingType EditingType, float X1, float Y1)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { throw new NotImplementedException(); }
        }

        public int Creator
        {
            get { throw new NotImplementedException(); }
        }

        public System.Collections.IEnumerator GetEnumerator()
        {
            return this.shapes.GetEnumerator();
        }

        public Microsoft.Office.Core.MsoTriState HasTitle
        {
            get { throw new NotImplementedException(); }
        }

        public dynamic Parent
        {
            get { throw new NotImplementedException(); }
        }

        public ShapeRange Paste()
        {
            throw new NotImplementedException();
        }

        public ShapeRange PasteSpecial(PpPasteDataType DataType = PpPasteDataType.ppPasteDefault, Microsoft.Office.Core.MsoTriState DisplayAsIcon = Microsoft.Office.Core.MsoTriState.msoFalse, string IconFileName = "", int IconIndex = 0, string IconLabel = "", Microsoft.Office.Core.MsoTriState Link = Microsoft.Office.Core.MsoTriState.msoFalse)
        {
            throw new NotImplementedException();
        }

        public Placeholders Placeholders
        {
            get { throw new NotImplementedException(); }
        }

        public ShapeRange Range([System.Runtime.InteropServices.OptionalAttribute]object Index)
        {
            throw new NotImplementedException();
        }

        public void SelectAll()
        {
            throw new NotImplementedException();
        }

        public Shape Title
        {
            get { throw new NotImplementedException(); }
        }

        public Shape this[object Index]
        {
            get { throw new NotImplementedException(); }
        }
    }
}
