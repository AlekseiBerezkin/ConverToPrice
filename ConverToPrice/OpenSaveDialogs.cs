// Decompiled with JetBrains decompiler
// Type: PriceConvert.OpenSaveDialogs
// Assembly: PriceConvert, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: DA1B6D49-AD90-4428-B938-61BD8BB8A206
// Assembly location: C:\MyProjects\HEOX.EBAY\book\PC\PriceConvert.exe

using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConverToPrice
{
  public static class OpenSaveDialogs
  {
    private static object lo = new object();

    public static List<string> OpenDialog(
      string settings_file,
      bool multi,
      DialogFilterType filter_type = DialogFilterType.AllFiles,
      string filter = "")
    {
      string path = "";
      lock (OpenSaveDialogs.lo)
      {
        if (File.Exists(settings_file))
        {
          using (StreamReader streamReader = new StreamReader(settings_file))
            path = streamReader.ReadLine();
        }
      }
      OpenFileDialog openFileDialog = new OpenFileDialog();
      if (Directory.Exists(path))
        openFileDialog.InitialDirectory = path;
      openFileDialog.FileName = "";
      openFileDialog.Multiselect = multi;
      if (filter == "")
      {
        switch (filter_type)
        {
          case DialogFilterType.Xlsx:
            openFileDialog.Filter = "Xlsx files (*.xlsx) | *.xlsx";
            break;
          case DialogFilterType.Xls:
            openFileDialog.Filter = "Xls files (*.xls) | *.xls";
            break;
          case DialogFilterType.XlsxXls:
            openFileDialog.Filter = "Xls/Xlsx files (*.xls, *.xlsx) | *.xls; *.xlsx";
            break;
          case DialogFilterType.Txt:
            openFileDialog.Filter = "Txt files (*.txt) | *.txt";
            break;
          case DialogFilterType.XlsxXlsTxt:
            openFileDialog.Filter = "Xls/Xlsx/Txt files (*.xls, *.xlsx, *.txt) | *.xls; *.xlsx; *.txt";
            break;
          case DialogFilterType.AllImages:
            openFileDialog.Filter = "Image files (*.jpg, *.jpeg, *.gif, *.png) | *.jpg; *.jpeg; *gif; *.png";
            break;
          case DialogFilterType.Pdf:
            openFileDialog.Filter = "Pdf files (*.pdf) | *.pdf";
            break;
          case DialogFilterType.HtmHtml:
            openFileDialog.Filter = "Html files (*.htm, *.html) | *.htm; *.html";
            break;
          case DialogFilterType.AllFiles:
            openFileDialog.Filter = "All files (*.*)|*.*";
            break;
        }
      }
      else
        openFileDialog.Filter = filter;
      bool? nullable = openFileDialog.ShowDialog();
      bool flag = true;
      if (!(nullable.GetValueOrDefault() == flag & nullable.HasValue))
        return (List<string>) null;
      string directoryName = new FileInfo(((IEnumerable<string>) openFileDialog.FileNames).First<string>()).DirectoryName;
      lock (OpenSaveDialogs.lo)
      {
        using (StreamWriter streamWriter = new StreamWriter(settings_file))
          streamWriter.WriteLine(directoryName);
      }
      return ((IEnumerable<string>) openFileDialog.FileNames).ToList<string>();
    }

    public static string SaveDialog(
      string settings_file,
      DialogFilterType filter_type = DialogFilterType.AllFiles,
      string default_fn = "",
      string filter = "")
    {
      string path = "";
      lock (OpenSaveDialogs.lo)
      {
        if (File.Exists(settings_file))
        {
          using (StreamReader streamReader = new StreamReader(settings_file))
            path = streamReader.ReadLine();
        }
      }
      SaveFileDialog saveFileDialog = new SaveFileDialog();
      if (Directory.Exists(path))
        saveFileDialog.InitialDirectory = path;
      saveFileDialog.FileName = default_fn;
      if (filter == "")
      {
        switch (filter_type)
        {
          case DialogFilterType.Xlsx:
            saveFileDialog.Filter = "Xlsx files (*.xlsx) | *.xlsx";
            break;
          case DialogFilterType.Xls:
            saveFileDialog.Filter = "Xls files (*.xls) | *.xls";
            break;
          case DialogFilterType.XlsxXls:
            saveFileDialog.Filter = "Xls/Xlsx files (*.xls, *.xlsx) | *.xls; *.xlsx";
            break;
          case DialogFilterType.Txt:
            saveFileDialog.Filter = "Txt files (*.txt) | *.txt";
            break;
          case DialogFilterType.XlsxXlsTxt:
            saveFileDialog.Filter = "Xls/Xlsx/Txt files (*.xls, *.xlsx, *.txt) | *.xls; *.xlsx; *.txt";
            break;
          case DialogFilterType.AllImages:
            saveFileDialog.Filter = "Image files (*.jpg, *.jpeg, *.gif, *.png, *.tif, *.tiff) | *.jpg; *.jpeg; *gif; *.png; *.tif; *.tiff";
            break;
          case DialogFilterType.Pdf:
            saveFileDialog.Filter = "Pdf files (*.pdf) | *.pdf";
            break;
          case DialogFilterType.HtmHtml:
            saveFileDialog.Filter = "Html files (*.htm, *.html) | *.htm; *.html";
            break;
          case DialogFilterType.AllFiles:
            saveFileDialog.Filter = "All files (*.*)|*.*";
            break;
        }
      }
      else
        saveFileDialog.Filter = filter;
      bool? nullable = saveFileDialog.ShowDialog();
      bool flag = true;
      if (!(nullable.GetValueOrDefault() == flag & nullable.HasValue))
        return (string) null;
      string directoryName = new FileInfo(saveFileDialog.FileName).DirectoryName;
      lock (OpenSaveDialogs.lo)
      {
        using (StreamWriter streamWriter = new StreamWriter(settings_file))
          streamWriter.WriteLine(directoryName);
      }
      return saveFileDialog.FileName;
    }
  }
    public enum DialogFilterType
    {
        Xlsx,
        Xls,
        XlsxXls,
        Txt,
        XlsxXlsTxt,
        AllImages,
        Pdf,
        HtmHtml,
        AllFiles,
    }

}
