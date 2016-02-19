using System.Runtime.InteropServices;
using System.Windows.Forms;
using AddinX.Ribbon.Contract;
using AddinX.Ribbon.Contract.Command;
using AddinX.Ribbon.Contract.Enums;
using AddinX.Ribbon.ExcelDna;

namespace Ribbon2
{
    [ComVisible(true)]
    public class AddinRibbon : RibbonFluent
    {
        private string inputTextValue;
        private const string SampleTab = "SampleTab";
        private const string InputText = "TextInputId";
        private const string OutputText = "TextOutputId";

        protected override void CreateFluentRibbon(IRibbonBuilder build)
        {
            build.CustomUi.AddNamespace("x", "AddinX.Addin.Sample")
                .Ribbon
                .ContextualTabs(tabs =>
                    tabs.AddTabSet(set => set.SetIdMso(TabSetId.TabSetDrawingTools)
                        .Tabs(tab => tab.AddTab("Sample").SetId("SampleContextId")
                            .Groups(g => g.AddGroup("Custom group").SetId("CustomGroupContextId")
                                .Items(d =>
                                {
                                    d.AddButton("Button 1").SetId("ContextTabButton1")
                                        .LargeSize().ImageMso("HappyFace");
                                    d.AddButton("Button 2").SetId("ContextTabButton2")
                                        .LargeSize().ImageMso("Bold");
                                })
                            )
                        )
                    )
                )
                .Tabs(c =>
                {
                    c.AddTab("Sample").SetIdQ("x", SampleTab)
                        .Groups(g =>
                        {
                            g.AddGroup("Custom Group 1").SetIdQ("x", "CustomGroupOne")
                                .Items(d =>
                                {
                                    d.AddEditbox("input").SetId(InputText)
                                        .ImagePath("one").MaxLength(12)
                                        .SizeString(9);

                                    d.AddLabelControl().SetId(OutputText);
                                });
                        });
                });
        }

        protected override void CreateRibbonCommand(IRibbonCommands cmds)
        {
            // Custom Group 
            cmds.AddButtonCommand("ContextTabButton1")
                .Action(() => MessageBox.Show("Happy"));

            cmds.AddButtonCommand("ContextTabButton2")
               .Action(() => MessageBox.Show("Bold!"));

            cmds.AddLabelCommand(OutputText).GetLabel(() => inputTextValue);
            cmds.AddEditBoxCommand(InputText).OnChange(value =>
            {
                inputTextValue = value;
                Ribbon.InvalidateControl(OutputText);
            })
            .GetText(() => "Text");
        }

        public override void OnClosing()
        {
            AddinContext.ExcelApp.DisposeChildInstances(true);
            AddinContext.ExcelApp = null;
        }

        public override void OnOpening()
        {
            
        }
    }
}