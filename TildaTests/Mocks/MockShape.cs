using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Tilda.Models;

namespace TildaTests.Mocks
{
    class MockShape : PowerPoint.Shape
    {
        private TextEffectFormat textEffect;
        private PowerPoint.TextFrame2 textFrame;
        private FillFormat fill;
        private float rotation = 0;
        private float top = 0;
        private float left = 0;
        private float width = 0;
        private float height = 0;
        private string name = "";
        public Microsoft.Office.Core.MsoShapeType type;
        private int id = 5;
        private int z = 0;
        private MsoTriState isVisible;

        public MockShape(Microsoft.Office.Core.MsoShapeType type = Microsoft.Office.Core.MsoShapeType.msoTextBox,int z = 1)
        {
            this.textEffect = new MockTextEffectFormat();
            this.textFrame = new MockTextFrame2();
            this.fill = new MockFillFormat();
            this.type = type; //default = tb
            this.z = 1;

            this.id = Int32.Parse(Settings.NextRandomValue().Split('-')[0]);
        }

        public FillFormat Fill
        {
            get { return this.fill; }
            set { this.fill = value; }
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

        public int Id
        {
            get { return this.id; }
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

        public TextEffectFormat TextEffect
        {
            get { return this.textEffect; }
            set { this.textEffect = value; }
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

        public PowerPoint.ActionSettings ActionSettings {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.Adjustments Adjustments {
            get { throw new NotImplementedException(); }
        }

        public string AlternativeText {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public PowerPoint.AnimationSettings AnimationSettings {
            get { throw new NotImplementedException(); }
        }

        public dynamic Application {
            get { throw new NotImplementedException(); }
        }

        public void Apply() {
            throw new NotImplementedException();
        }

        public void ApplyAnimation() {
            throw new NotImplementedException();
        }

        public MsoAutoShapeType AutoShapeType {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoBackgroundStyleIndex BackgroundStyle {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoBlackWhiteMode BlackWhiteMode {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public PowerPoint.CalloutFormat Callout {
            get { throw new NotImplementedException(); }
        }

        public void CanvasCropBottom(float Increment) {
            throw new NotImplementedException();
        }

        public void CanvasCropLeft(float Increment) {
            throw new NotImplementedException();
        }

        public void CanvasCropRight(float Increment) {
            throw new NotImplementedException();
        }

        public void CanvasCropTop(float Increment) {
            throw new NotImplementedException();
        }

        public PowerPoint.CanvasShapes CanvasItems {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.Chart Chart {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState Child {
            get { throw new NotImplementedException(); }
        }

        public int ConnectionSiteCount {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState Connector {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.ConnectorFormat ConnectorFormat {
            get { throw new NotImplementedException(); }
        }

        public void ConvertTextToSmartArt(SmartArtLayout Layout) {
            throw new NotImplementedException();
        }

        public void Copy() {
            throw new NotImplementedException();
        }

        public int Creator {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.CustomerData CustomerData {
            get { throw new NotImplementedException(); }
        }

        public void Cut() {
            throw new NotImplementedException();
        }

        public void Delete() {
            throw new NotImplementedException();
        }

        public PowerPoint.Diagram Diagram {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.DiagramNode DiagramNode {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.ShapeRange Duplicate() {
            throw new NotImplementedException();
        }

        public void Export(string PathName, PowerPoint.PpShapeFormat Filter, int ScaleWidth = 0, int ScaleHeight = 0, PowerPoint.PpExportMode ExportMode = PowerPoint.PpExportMode.ppRelativeToSlide) {
        
        }

        PowerPoint.FillFormat PowerPoint.Shape.Fill {
            get { throw new NotImplementedException(); }
        }

        public void Flip(MsoFlipCmd FlipCmd) {
            throw new NotImplementedException();
        }

        public GlowFormat Glow {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.GroupShapes GroupItems {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState HasChart {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState HasDiagram {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState HasDiagramNode {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState HasSmartArt {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState HasTable {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState HasTextFrame {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState HorizontalFlip {
            get { throw new NotImplementedException(); }
        }

        public void IncrementLeft(float Increment) {
            throw new NotImplementedException();
        }

        public void IncrementRotation(float Increment) {
            throw new NotImplementedException();
        }

        public void IncrementTop(float Increment) {
            throw new NotImplementedException();
        }

        public PowerPoint.LineFormat Line {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.LinkFormat LinkFormat {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState LockAspectRatio {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public PowerPoint.MediaFormat MediaFormat {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.PpMediaType MediaType {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.ShapeNodes Nodes {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.OLEFormat OLEFormat {
            get { throw new NotImplementedException(); }
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.Shape ParentGroup {
            get { throw new NotImplementedException(); }
        }

        public void PickUp() {
            throw new NotImplementedException();
        }

        public void PickupAnimation() {
            throw new NotImplementedException();
        }

        public PowerPoint.PictureFormat PictureFormat {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.PlaceholderFormat PlaceholderFormat {
            get { throw new NotImplementedException(); }
        }

        public string RTF {
            set { throw new NotImplementedException(); }
        }

        public ReflectionFormat Reflection {
            get { throw new NotImplementedException(); }
        }

        public void RerouteConnections() {
            throw new NotImplementedException();
        }

        public void ScaleHeight(float Factor, MsoTriState RelativeToOriginalSize, MsoScaleFrom fScale = MsoScaleFrom.msoScaleFromTopLeft) {
            throw new NotImplementedException();
        }

        public void ScaleWidth(float Factor, MsoTriState RelativeToOriginalSize, MsoScaleFrom fScale = MsoScaleFrom.msoScaleFromTopLeft) {
            throw new NotImplementedException();
        }

        public Script Script {
            get { throw new NotImplementedException(); }
        }

        public void Select(MsoTriState Replace = MsoTriState.msoTrue) {
            throw new NotImplementedException();
        }

        public void SetShapesDefaultProperties() {
            throw new NotImplementedException();
        }

        public PowerPoint.ShadowFormat Shadow {
            get { throw new NotImplementedException(); }
        }

        public MsoShapeStyleIndex ShapeStyle {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public SmartArt SmartArt {
            get { throw new NotImplementedException(); }
        }

        public SoftEdgeFormat SoftEdge {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.SoundFormat SoundFormat {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.Table Table {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.Tags Tags {
            get { throw new NotImplementedException(); }
        }

        PowerPoint.TextEffectFormat PowerPoint.Shape.TextEffect {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.TextFrame TextFrame {
            get { throw new NotImplementedException(); }
        }

        PowerPoint.TextFrame2 PowerPoint.Shape.TextFrame2 {
            get { return this.textFrame; }
        }

        public PowerPoint.ThreeDFormat ThreeD {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.ShapeRange Ungroup() {
            throw new NotImplementedException();
        }

        public void UpgradeMedia() {
            throw new NotImplementedException();
        }

        public MsoTriState VerticalFlip {
            get { throw new NotImplementedException(); }
        }

        public dynamic Vertices {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState Visible {
            get {
                return this.isVisible;
            }
            set {
                this.isVisible = value;
            }
        }

        public void ZOrder(MsoZOrderCmd ZOrderCmd) {
            if(ZOrderCmd == MsoZOrderCmd.msoSendToBack)
                this.z = 0;
        }

        public int ZOrderPosition {
            get { return this.z; }
        }
    }
}
