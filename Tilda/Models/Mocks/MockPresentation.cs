using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace Tilda.Models {
    class MockPresentation : Presentation {

        public void AcceptAll() {
            throw new NotImplementedException();
        }

        public void AddBaseline(string FileName = "") {
            throw new NotImplementedException();
        }

        public Master AddTitleMaster() {
            throw new NotImplementedException();
        }

        public void AddToFavorites() {
            throw new NotImplementedException();
        }

        public Application Application {
            get { throw new NotImplementedException(); }
        }

        public void ApplyTemplate(string FileName) {
            throw new NotImplementedException();
        }

        public void ApplyTheme(string themeName) {
            throw new NotImplementedException();
        }

        public Broadcast Broadcast {
            get { throw new NotImplementedException(); }
        }

        public dynamic BuiltInDocumentProperties {
            get { throw new NotImplementedException(); }
        }

        public bool CanCheckIn() {
            throw new NotImplementedException();
        }

        public void CheckIn(bool SaveChanges,
            [System.Runtime.InteropServices.OptionalAttribute]object Comments , 
            [System.Runtime.InteropServices.OptionalAttribute]object MakePublic) {
            throw new NotImplementedException();
        }

        public void CheckInWithVersion(bool SaveChanges, [System.Runtime.InteropServices.OptionalAttribute]object Comments, [System.Runtime.InteropServices.OptionalAttribute]object MakePublic, [System.Runtime.InteropServices.OptionalAttribute]object VersionType) {
            throw new NotImplementedException();
        }

        public void Close() {
            throw new NotImplementedException();
        }

        public Coauthoring Coauthoring {
            get { throw new NotImplementedException(); }
        }

        public ColorSchemes ColorSchemes {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.CommandBars CommandBars {
            get { throw new NotImplementedException(); }
        }

        public dynamic Container {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MetaProperties ContentTypeProperties {
            get { throw new NotImplementedException(); }
        }

        public void Convert() {
            throw new NotImplementedException();
        }

        public void Convert2(string FileName) {
            throw new NotImplementedException();
        }

        public void CreateVideo(string FileName, bool UseTimingsAndNarrations = true, int DefaultSlideDuration = 5, int VertResolution = 720, int FramesPerSecond = 30, int Quality = 85) {
            throw new NotImplementedException();
        }

        public PpMediaTaskStatus CreateVideoStatus {
            get { throw new NotImplementedException(); }
        }

        public dynamic CustomDocumentProperties {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.CustomXMLParts CustomXMLParts {
            get { throw new NotImplementedException(); }
        }

        public CustomerData CustomerData {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoLanguageID DefaultLanguageID {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Shape DefaultShape {
            get { throw new NotImplementedException(); }
        }

        public void DeleteSection(int Index) {
            throw new NotImplementedException();
        }

        public Designs Designs {
            get { throw new NotImplementedException(); }
        }

        public void DisableSections() {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState DisplayComments {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.DocumentInspectors DocumentInspectors {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.DocumentLibraryVersions DocumentLibraryVersions {
            get { throw new NotImplementedException(); }
        }

        public string EncryptionProvider {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public void EndReview() {
            throw new NotImplementedException();
        }

        public void EnsureAllMediaUpgraded() {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState EnvelopeVisible {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public void Export(string Path, string FilterName, int ScaleWidth = 0, int ScaleHeight = 0) {
            throw new NotImplementedException();
        }

        public void ExportAsFixedFormat(string Path, PpFixedFormatType FixedFormatType, PpFixedFormatIntent Intent, Microsoft.Office.Core.MsoTriState FrameSlides, PpPrintHandoutOrder HandoutOrder, PpPrintOutputType OutputType, Microsoft.Office.Core.MsoTriState PrintHiddenSlides, PrintRange PrintRange, PpPrintRangeType RangeType, string SlideShowName, bool IncludeDocProperties, bool KeepIRMSettings, bool DocStructureTags, bool BitmapMissingFonts, bool UseISO19005_1, [System.Runtime.InteropServices.OptionalAttribute]object ExternalExporter) {
            throw new NotImplementedException();
        }

        public ExtraColors ExtraColors {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoFarEastLineBreakLanguageID FarEastLineBreakLanguage {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public PpFarEastLineBreakLevel FarEastLineBreakLevel {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public bool Final {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public void FollowHyperlink(string Address, string SubAddress = "", bool NewWindow = false, bool AddHistory = true, string ExtraInfo = "", Microsoft.Office.Core.MsoExtraInfoMethod Method = Microsoft.Office.Core.MsoExtraInfoMethod.msoMethodGet, string HeaderInfo = "") {
            throw new NotImplementedException();
        }

        public Fonts Fonts {
            get { throw new NotImplementedException(); }
        }

        public string FullName {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.WorkflowTasks GetWorkflowTasks() {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.WorkflowTemplates GetWorkflowTemplates() {
            throw new NotImplementedException();
        }

        public float GridDistance {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.HTMLProject HTMLProject {
            get { throw new NotImplementedException(); }
        }

        public Master HandoutMaster {
            get { throw new NotImplementedException(); }
        }

        public bool HasHandoutMaster {
            get { throw new NotImplementedException(); }
        }

        public bool HasNotesMaster {
            get { throw new NotImplementedException(); }
        }

        public PpRevisionInfo HasRevisionInfo {
            get { throw new NotImplementedException(); }
        }

        public bool HasSections {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState HasTitleMaster {
            get { throw new NotImplementedException(); }
        }

        public bool HasVBProject {
            get { throw new NotImplementedException(); }
        }

        public bool InMergeMode {
            get { throw new NotImplementedException(); }
        }

        public PpDirection LayoutDirection {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public void LockServerFile() {
            throw new NotImplementedException();
        }

        public void MakeIntoTemplate(Microsoft.Office.Core.MsoTriState IsDesignTemplate) {
            throw new NotImplementedException();
        }

        public void Merge(string Path) {
            throw new NotImplementedException();
        }

        public void MergeWithBaseline(string withPresentation, string baselinePresentation) {
            throw new NotImplementedException();
        }

        public string Name {
            get { throw new NotImplementedException(); }
        }

        public void NewSectionAfter(int Index, bool AfterSlide, string sectionTitle, out int newSectionIndex) {
            throw new NotImplementedException();
        }

        public DocumentWindow NewWindow() {
            throw new NotImplementedException();
        }

        public string NoLineBreakAfter {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public string NoLineBreakBefore {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Master NotesMaster {
            get { throw new NotImplementedException(); }
        }

        public PageSetup PageSetup {
            get { return new MockPageSetup(); }
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public string Password {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public string PasswordEncryptionAlgorithm {
            get { throw new NotImplementedException(); }
        }

        public bool PasswordEncryptionFileProperties {
            get { throw new NotImplementedException(); }
        }

        public int PasswordEncryptionKeyLength {
            get { throw new NotImplementedException(); }
        }

        public string PasswordEncryptionProvider {
            get { throw new NotImplementedException(); }
        }

        public string Path {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.Permission Permission {
            get { throw new NotImplementedException(); }
        }

        public PrintOptions PrintOptions {
            get { throw new NotImplementedException(); }
        }

        public void PrintOut(int From, int To, string PrintToFile, int Copies, Microsoft.Office.Core.MsoTriState Collate) {
            throw new NotImplementedException();
        }

        public PublishObjects PublishObjects {
            get { throw new NotImplementedException(); }
        }

        public void PublishSlides(string SlideLibraryUrl, bool Overwrite = false, bool UseSlideOrder = false) {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState ReadOnly {
            get { throw new NotImplementedException(); }
        }

        public void RejectAll() {
            throw new NotImplementedException();
        }

        public void ReloadAs(Microsoft.Office.Core.MsoEncoding cp) {
            throw new NotImplementedException();
        }

        public void RemoveBaseline() {
            throw new NotImplementedException();
        }

        public void RemoveDocumentInformation(PpRemoveDocInfoType Type) {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState RemovePersonalInformation {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public void ReplyWithChanges(bool ShowMessage = true) {
            throw new NotImplementedException();
        }

        public Research Research {
            get { throw new NotImplementedException(); }
        }

        public void Save() {
            throw new NotImplementedException();
        }

        public void SaveAs(string FileName, PpSaveAsFileType FileFormat = PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState EmbedTrueTypeFonts = Microsoft.Office.Core.MsoTriState.msoTriStateMixed) {
            throw new NotImplementedException();
        }

        public void SaveCopyAs(string FileName, PpSaveAsFileType FileFormat = PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState EmbedTrueTypeFonts = Microsoft.Office.Core.MsoTriState.msoTriStateMixed) {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState Saved {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public int SectionCount {
            get { throw new NotImplementedException(); }
        }

        public SectionProperties SectionProperties {
            get { throw new NotImplementedException(); }
        }

        public void SendFaxOverInternet(string Recipients = "", string Subject = "", bool ShowMessage = false) {
            throw new NotImplementedException();
        }

        public void SendForReview(string Recipients, string Subject, bool ShowMessage, [System.Runtime.InteropServices.OptionalAttribute]object IncludeAttachment) {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.ServerPolicy ServerPolicy {
            get { throw new NotImplementedException(); }
        }

        public void SetPasswordEncryptionOptions(string PasswordEncryptionProvider, string PasswordEncryptionAlgorithm, int PasswordEncryptionKeyLength, bool PasswordEncryptionFileProperties) {
            throw new NotImplementedException();
        }

        public void SetUndoText(string Text) {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.SharedWorkspace SharedWorkspace {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.SignatureSet Signatures {
            get { throw new NotImplementedException(); }
        }

        public Master SlideMaster {
            get { throw new NotImplementedException(); }
        }

        public SlideShowSettings SlideShowSettings {
            get { throw new NotImplementedException(); }
        }

        public SlideShowWindow SlideShowWindow {
            get { throw new NotImplementedException(); }
        }

        public Slides Slides {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState SnapToGrid {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.Sync Sync {
            get { throw new NotImplementedException(); }
        }

        public Tags Tags {
            get { throw new NotImplementedException(); }
        }

        public string TemplateName {
            get { throw new NotImplementedException(); }
        }

        public Master TitleMaster {
            get { throw new NotImplementedException(); }
        }

        public void Unused() {
            throw new NotImplementedException();
        }

        public void UpdateLinks() {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState VBASigned {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Vbe.Interop.VBProject VBProject {
            get { throw new NotImplementedException(); }
        }

        public WebOptions WebOptions {
            get { throw new NotImplementedException(); }
        }

        public void WebPagePreview() {
            throw new NotImplementedException();
        }

        public DocumentWindows Windows {
            get { throw new NotImplementedException(); }
        }

        public string WritePassword {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public void sblt(string s) {
            throw new NotImplementedException();
        }

        public string sectionTitle(int Index) {
            throw new NotImplementedException();
        }
    }
}
