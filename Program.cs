
using Auxi_PowerPointEdit.Helpers;
using DocumentFormat.OpenXml.Packaging;
using System;

namespace Auxi_PowerPointEdit
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                string pptPath = String.Format("D:\\Source Code\\AUXI_EditPowerPoint\\Auxi_PowerPointEdit\\Assets\\auxi C# Interview-1.pptx");
                PPTHelper pPTHelper = new PPTHelper();
                pPTHelper.OpenPptxFile(pptPath);
                SlidePart OutputSlide = pPTHelper.CloneInputSlide();
                PPTHelper.FormatTitle(OutputSlide);
                PPTHelper.FormatFLow(OutputSlide);
                PPTHelper.FormatBlPoint(OutputSlide);
                pPTHelper.Dispose();

            }
            finally
            {
            }
        }

    }
}
