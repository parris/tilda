using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks
{
    class MockShape : Shape
    {
        private TextEffectFormat textEffect;
        private FillFormat fill;
        private float rotation = 0;
        private float top = 0;
        private float left = 0;
        private float width = 0;
        private float height = 0;
        private string name = "";
        public Microsoft.Office.Core.MsoShapeType type;
        private int id = 5;

        public MockShape(Microsoft.Office.Core.MsoShapeType type = Microsoft.Office.Core.MsoShapeType.msoTextBox)
        {
            this.textEffect = new MockTextEffectFormat();
            this.fill = new MockFillFormat();
            this.type = type; //default = tb
        }

        public ActionSettings ActionSettings
        {
            get { throw new NotImplementedException(); }
        }

        public Adjustments Adjustments
        {
            get { throw new NotImplementedException(); }
        }

        public string AlternativeText
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public AnimationSettings AnimationSettings
        {
            get { throw new NotImplementedException(); }
        }

        public dynamic Application
        {
            get { throw new NotImplementedException(); }
        }

        public void Apply()
        {
            throw new NotImplementedException();
        }

        public void ApplyAnimation()
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoAutoShapeType AutoShapeType
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoBackgroundStyleIndex BackgroundStyle
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoBlackWhiteMode BlackWhiteMode
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public CalloutFormat Callout
        {
            get { throw new NotImplementedException(); }
        }

        public void CanvasCropBottom(float Increment)
        {
            throw new NotImplementedException();
        }

        public void CanvasCropLeft(float Increment)
        {
            throw new NotImplementedException();
        }

        public void CanvasCropRight(float Increment)
        {
            throw new NotImplementedException();
        }

        public void CanvasCropTop(float Increment)
        {
            throw new NotImplementedException();
        }

        public CanvasShapes CanvasItems
        {
            get { throw new NotImplementedException(); }
        }

        public Chart Chart
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState Child
        {
            get { throw new NotImplementedException(); }
        }

        public int ConnectionSiteCount
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState Connector
        {
            get { throw new NotImplementedException(); }
        }

        public ConnectorFormat ConnectorFormat
        {
            get { throw new NotImplementedException(); }
        }

        public void ConvertTextToSmartArt(Microsoft.Office.Core.SmartArtLayout Layout)
        {
            throw new NotImplementedException();
        }

        public void Copy()
        {
            throw new NotImplementedException();
        }

        public int Creator
        {
            get { throw new NotImplementedException(); }
        }

        public CustomerData CustomerData
        {
            get { throw new NotImplementedException(); }
        }

        public void Cut()
        {
            throw new NotImplementedException();
        }

        public void Delete()
        {
            throw new NotImplementedException();
        }

        public Diagram Diagram
        {
            get { throw new NotImplementedException(); }
        }

        public DiagramNode DiagramNode
        {
            get { throw new NotImplementedException(); }
        }

        public ShapeRange Duplicate()
        {
            throw new NotImplementedException();
        }

        public void Export(string PathName, PpShapeFormat Filter, int ScaleWidth = 0, int ScaleHeight = 0, PpExportMode ExportMode = PpExportMode.ppRelativeToSlide)
        {
            
        }

        public FillFormat Fill
        {
            get { return this.fill; }
            set { this.fill = value; }
        }

        public void Flip(Microsoft.Office.Core.MsoFlipCmd FlipCmd)
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.GlowFormat Glow
        {
            get { throw new NotImplementedException(); }
        }

        public GroupShapes GroupItems
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState HasChart
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState HasDiagram
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState HasDiagramNode
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState HasSmartArt
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState HasTable
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState HasTextFrame
        {
            get { throw new NotImplementedException(); }
        }

        public float Height
        {
            get
            {
                return this.height;
            }
            set
            {
                this.height = value;
            }
        }

        public Microsoft.Office.Core.MsoTriState HorizontalFlip
        {
            get { throw new NotImplementedException(); }
        }

        public int Id
        {
            get { return this.id; }
        }

        public void IncrementLeft(float Increment)
        {
            throw new NotImplementedException();
        }

        public void IncrementRotation(float Increment)
        {
            throw new NotImplementedException();
        }

        public void IncrementTop(float Increment)
        {
            throw new NotImplementedException();
        }

        public float Left
        {
            get
            {
                return this.left;
            }
            set
            {
                this.left = value;
            }
        }

        public LineFormat Line
        {
            get { throw new NotImplementedException(); }
        }

        public LinkFormat LinkFormat
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState LockAspectRatio
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public MediaFormat MediaFormat
        {
            get { throw new NotImplementedException(); }
        }

        public PpMediaType MediaType
        {
            get { throw new NotImplementedException(); }
        }

        public string Name
        {
            get
            {
                return this.name;
            }
            set
            {
                this.name = value;
            }
        }

        public ShapeNodes Nodes
        {
            get { throw new NotImplementedException(); }
        }

        public OLEFormat OLEFormat
        {
            get { throw new NotImplementedException(); }
        }

        public dynamic Parent
        {
            get { throw new NotImplementedException(); }
        }

        public Shape ParentGroup
        {
            get { throw new NotImplementedException(); }
        }

        public void PickUp()
        {
            throw new NotImplementedException();
        }

        public void PickupAnimation()
        {
            throw new NotImplementedException();
        }

        public PictureFormat PictureFormat
        {
            get { throw new NotImplementedException(); }
        }

        public PlaceholderFormat PlaceholderFormat
        {
            get { throw new NotImplementedException(); }
        }

        public string RTF
        {
            set { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.ReflectionFormat Reflection
        {
            get { throw new NotImplementedException(); }
        }

        public void RerouteConnections()
        {
            throw new NotImplementedException();
        }

        public float Rotation
        {
            get
            {
                return this.rotation;
            }
            set
            {
                this.rotation = value;
            }
        }

        public void ScaleHeight(float Factor, Microsoft.Office.Core.MsoTriState RelativeToOriginalSize, Microsoft.Office.Core.MsoScaleFrom fScale = Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)
        {
            throw new NotImplementedException();
        }

        public void ScaleWidth(float Factor, Microsoft.Office.Core.MsoTriState RelativeToOriginalSize, Microsoft.Office.Core.MsoScaleFrom fScale = Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.Script Script
        {
            get { throw new NotImplementedException(); }
        }

        public void Select(Microsoft.Office.Core.MsoTriState Replace = Microsoft.Office.Core.MsoTriState.msoTrue)
        {
            throw new NotImplementedException();
        }

        public void SetShapesDefaultProperties()
        {
            throw new NotImplementedException();
        }

        public ShadowFormat Shadow
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoShapeStyleIndex ShapeStyle
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.SmartArt SmartArt
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.SoftEdgeFormat SoftEdge
        {
            get { throw new NotImplementedException(); }
        }

        public SoundFormat SoundFormat
        {
            get { throw new NotImplementedException(); }
        }

        public Table Table
        {
            get { throw new NotImplementedException(); }
        }

        public Tags Tags
        {
            get { throw new NotImplementedException(); }
        }

        public TextEffectFormat TextEffect
        {
            get { return this.textEffect; }
            set { this.textEffect = value; }
        }

        public TextFrame TextFrame
        {
            get { throw new NotImplementedException(); }
        }

        public TextFrame2 TextFrame2
        {
            get { throw new NotImplementedException(); }
        }

        public ThreeDFormat ThreeD
        {
            get { throw new NotImplementedException(); }
        }

        public string Title
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public float Top
        {
            get
            {
                return this.top;
            }
            set
            {
                this.top = value;
            }
        }

        public Microsoft.Office.Core.MsoShapeType Type
        {
            get { return this.type; }
        }

        public ShapeRange Ungroup()
        {
            throw new NotImplementedException();
        }

        public void UpgradeMedia()
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState VerticalFlip
        {
            get { throw new NotImplementedException(); }
        }

        public dynamic Vertices
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState Visible
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public float Width
        {
            get
            {
                return this.width;
            }
            set
            {
                this.width = value;
            }
        }

        public void ZOrder(Microsoft.Office.Core.MsoZOrderCmd ZOrderCmd)
        {
            throw new NotImplementedException();
        }

        public int ZOrderPosition
        {
            get { throw new NotImplementedException(); }
        }
    }
}
