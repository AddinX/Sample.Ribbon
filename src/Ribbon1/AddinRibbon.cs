using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AddinX.Ribbon.Contract;
using AddinX.Ribbon.Contract.Command;
using AddinX.Ribbon.ExcelDna;
using Sample.AddIn.Ribbon.Data;
using Sample.AddIn.Ribbon.Utils;

namespace Sample.AddIn.Ribbon
{
    [ComVisible(true)]
    public class AddinRibbon : RibbonFluent
    {
        private const string SampleTab = "SampleTab";
        private const string OtherButtonBox = "otherButtonBox";
        private const string BoxLauncherId = "BoxLauncher";
        private const string ButtonExtra = "ButtonExtra";

        private const string HappyButtonId = "HappyButton1";
        private const string CheckboxDynamicGallery = "CheckboxDynamicGallery";
        private const string OptionId = "OptionId";
        private const string DynamicGalleryId = "DynamicGalleryId";
        
        private const string ComboBoxId = "comboboxId";
        private const string DropDownId = "dropDownId";


        private ListItems content;
        private bool checkboxPressed;

        public override Bitmap OnLoadImage(string imageName)
        {
            switch (imageName)
            {
                case "desert":
                    return Properties.Resources.Desert;
                case "forest":
                    return Properties.Resources.Forest;
                case "toucan":
                    return Properties.Resources.Toucan;
                case "tree":
                    return Properties.Resources.Tree;
            }
            return Properties.Resources.one;
        }

        protected override void CreateFluentRibbon(IRibbonBuilder builder)
        {
            builder.CustomUi.AddNamespace("x","AddinX.Addin.Sample")
                .Ribbon.Tabs(c =>
            {
                c.AddTab("Sample").SetIdQ("x",SampleTab)
                    .Groups(g =>
                    {
                        g.AddGroup("Custom Group 1").SetIdQ("x","CustomGroupOne")
                            .Items(d =>
                            {
                                d.AddButton("Button 1")
                                    .SetId("button1")
                                    .LargeSize()
                                    .ImageMso("Repeat");

                                d.AddBox().SetId(OtherButtonBox)
                                    .HorizontalDisplay()
                                    .AddItems(i =>
                                    {
                                        i.AddButton("Button 2").SetId("button2")
                                            .NormalSize().NoImage().ShowLabel()
                                            .Screentip("Special Button")
                                            .Supertip("Display special content :)");

                                        i.AddButton("Button 3")
                                           .SetId("button3")
                                           .NormalSize()
                                           .ImageMso("Bold");
                                    });
                                    
                               d.AddGallery("Gallery").SetId("galleryOneId")
                                   .ShowLabel().NormalSize().ImageMso("HappyFace")
                                   .HideItemLabel().ShowItemImage()
                                   .AddItems(v =>
                                   {
                                       v.AddItem("Desert").SetId("Item1").ImagePath("desert");
                                       v.AddItem("Forest").SetId("Item2").ImagePath("forest");
                                       v.AddItem("Toucan").SetId("Item3").ImagePath("toucan");
                                       v.AddItem("Tree").SetId("Item4").ImagePath("tree");
                                   }).AddButtons(b => b.AddButton("Extra...")
                                                        .SetId(ButtonExtra))
                                    .ItemHeight(88).ItemWidth(68);

                                d.AddBox().SetId("BoxIdMsoControls")
                                .HorizontalDisplay().AddItems(b =>
                                {
                                    b.AddButton("Save File").SetIdMso("FileSave")
                                        .NormalSize().ImageMso("FileSave");
                                    b.AddButton("New").SetIdMso("FileNew")
                                        .NormalSize().ImageMso("FileNew");
                                    b.AddButton("Print").SetIdMso("PrintTitles")
                                        .NormalSize().ImageMso("PrintTitles")
                                        .HideLabel()
                                        .Screentip("Print Titles!")
                                        .Supertip("Print the titles of the current sheet");
                                });
                                
                            })
                            .DialogBoxLauncher(i => i.AddDialogBoxLauncher()
                                                .SetId(BoxLauncherId).Supertip("Box launcher")); ;

                        g.AddGroup("Custom Group 2").SetId("CustomGroupTwo")
                           .Items(d =>
                           {
                               d.AddMenu("Menu").SetId(OptionId).ShowLabel()
                                   .ImageMso("FileSendMenu").LargeSize()
                                   .ItemLargeSize().AddItems(
                                       v =>
                                       {
                                           v.AddCheckbox("Enable Dynamic gallery")
                                                .SetId(CheckboxDynamicGallery);
                                           v.AddSeparator("Extra");
                                           v.AddButton("Happy")
                                               .SetId(HappyButtonId)
                                               .ImageMso("HappyFace");
                                           v.AddGallery("Dynamic Option")
                                               .SetId(DynamicGalleryId)
                                               .ShowLabel().NoImage()
                                               .HideItemLabel().ShowItemImage()
                                               .DynamicItems().NumberRows(6)
                                               .NumberColumns(1);
                                       });
                               d.AddComboBox("Items 1")
                                .SetId(ComboBoxId)
                                .ShowLabel().NoImage()
                                .DynamicItems();

                               d.AddDropDown("Items 2")
                                    .SetId(DropDownId)
                                    .ShowLabel().NoImage()
                                    .ShowItemLabel().ShowItemImage()
                                    .DynamicItems();

                           });
                    });
            });
        }

        protected override void CreateRibbonCommand(IRibbonCommands cmds)
        {
            // Custom Group 1
            cmds.AddButtonCommand("button3")
                .IsEnabled(() => AddinContext.ExcelApp.Worksheets.Count() > 2)
                .Action(() => MessageBox.Show("button 3 clicked"));

            cmds.AddButtonCommand("button2")
                .Action(() => MessageBox.Show("To enable the 'Button 3' you need to have 3 sheets."));

            cmds.AddButtonCommand("button1")
                .Action(() => MessageBox.Show("You need two sheet so see the others buttons"));

            cmds.AddBoxCommand(OtherButtonBox)
                .IsVisible(() => AddinContext.ExcelApp.Worksheets.Count() > 1);

            cmds.AddButtonCommand(ButtonExtra).Action(() => MessageBox.Show("More..."));

            cmds.AddDialogBoxLauncherCommand(BoxLauncherId)
               .Action(() => MessageBox.Show("Dialog Box clicked"));

            cmds.AddGalleryCommand("galleryOneId")
                    .Action(index =>
                    {
                        switch (index)
                        {
                            case 0:
                                MessageBox.Show("Desert");
                                break;
                            case 1:
                                MessageBox.Show("Forest");
                                break;
                            case 2:
                                MessageBox.Show("Toucan");
                                break;
                            case 3:
                                MessageBox.Show("Tree");
                                break;
                        }
                    });

            // Custom Group 3
            cmds.AddButtonCommand(HappyButtonId).Action(() => MessageBox.Show("Be Happy !!!"));

            cmds.AddCheckBoxCommand(CheckboxDynamicGallery).Action(isPressed =>
            {
                checkboxPressed = isPressed;
                Ribbon.InvalidateControl(DynamicGalleryId);
                MessageBox.Show(isPressed ? "Show number pressed" : "Show number NOT pressed");
            }).Pressed(() => false);

            cmds.AddGalleryCommand(DynamicGalleryId)
                .IsEnabled(() => checkboxPressed)
                .ItemCounts(content.Count)
                .ItemsId(content.Ids)
                .ItemsLabel(content.Labels)
                .ItemsImage(() => content.Images())
                .ItemsSupertip(content.SuperTips)
                .Action(i => MessageBox.Show(@"You selected: " + (i + 1)));


            cmds.AddDropDownCommand(DropDownId)
               .ItemCounts(content.Count)
               .ItemsId(content.Ids)
               .ItemsLabel(content.Labels)
               .ItemsImage(() => content.Images())
               .ItemsSupertip(content.SuperTips)
               .Action(i => MessageBox.Show(@"You selected:" + (i + 1)));

            cmds.AddComboBoxCommand(ComboBoxId)
                .ItemCounts(content.Count)
                .ItemsId(content.Ids)
                .ItemsLabel(content.Labels)
                .ItemsSupertip(content.SuperTips)
                .GetText(() => "Numbers")
                .OnChange((value) => MessageBox.Show(@"You selected:" + value));
        }

        public override void OnClosing()
        {
            AddinContext.ExcelApp.DisposeChildInstances(true);
            AddinContext.ExcelApp = null;
        }

        public override void OnOpening()
        {
            AddinContext.ExcelApp.SheetActivateEvent += (e) => RefreshRibbon();
            AddinContext.ExcelApp.SheetChangeEvent += (a, e) => RefreshRibbon();
            
            LoadContent();
        }

        private void LoadContent()
        {
            content = new ListItems();
            content.Add(new SingleItem
            {
                Label = "First Item"
                ,
                SuperTip = "The first Item"
                ,
                Image = ResizeImage.Resize(Properties.Resources.one, 32, 32)
            });
            content.Add(new SingleItem
            {
                Label = "Second Item"
                ,
                SuperTip = "The second Item"
                ,
                Image = ResizeImage.Resize(Properties.Resources.two, 32, 32)
            });
            content.Add(new SingleItem
            {
                Label = "Third Item"
                ,
                SuperTip = "The third Item"
                ,
                Image = ResizeImage.Resize(Properties.Resources.three, 32, 32)
            });
            content.Add(new SingleItem
            {
                Label = "Fourth Item"
                ,
                SuperTip = "The fourth Item"
                ,
                Image = ResizeImage.Resize(Properties.Resources.four, 32, 32)
            });
            content.Add(new SingleItem
            {
                Label = "Fifth Item"
                ,
                SuperTip = "The fifth Item"
                ,
                Image = ResizeImage.Resize(Properties.Resources.five, 32, 32)
            });
        }

        private void RefreshRibbon()
        {
            Ribbon?.Invalidate();
        }
    }
}