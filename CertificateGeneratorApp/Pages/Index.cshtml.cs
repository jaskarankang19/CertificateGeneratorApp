using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using System.IO;
using Xceed.Words.NET;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using DocumentFormat.OpenXml;
using System.Diagnostics;
//using System.Reflection.Metadata;

public class IndexModel : PageModel
{


    public void OnGet()
    {

    }
    public IActionResult OnPost()
    {

        return Page();
    }


}
