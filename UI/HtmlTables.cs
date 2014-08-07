using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.UI.HtmlControls;
using System.Web.UI;

namespace UI
{
    public class HtmlTables : Control
    {
        public List<Sheet> Sheets;

        public HtmlTables(List<Sheet> sheets)
        {
            Sheets = sheets;
            foreach (var sheet in Sheets)
            {
                var table = new HtmlGenericControl("table");
                Controls.Add(table);
                var colgroup = new HtmlGenericControl("colgroup");
                var tbody = new HtmlGenericControl("tbody");

                //add table children
                table.Controls.Add(colgroup);
                table.Controls.Add(tbody);

                //set title
                if (!string.IsNullOrEmpty(sheet.title))
                {
                    table.Attributes["title"] = sheet.title;
                }

                //read metadata
                var metadata = sheet.metadata;
                if (metadata != null)
                {
                    //make frozenAt
                    var frozenAt = metadata.frozenAt;
                    if (frozenAt != null)
                    {
                        if (frozenAt.col > 0)
                        {
                            table.Attributes["data-frozenAtCol"] = frozenAt.col.ToString();
                        }
                        if (frozenAt.row > 0)
                        {
                            table.Attributes["data-frozenAtRow"] = frozenAt.row.ToString();
                        }
                    }

                    //create widths
                    foreach(var width in metadata.widths)
                    {
                        var col = new HtmlGenericControl("col");
                        col.Style["width"] = width.Replace("px", "") + "px";
                        colgroup.Controls.Add(col);
                    }
                }

                foreach (var row in sheet.rows)
                {
                    var tr = new HtmlGenericControl("tr");
                    tbody.Controls.Add(tr);
                    if (row.height > 0)
                    {
                        tr.Style["height"] = row.height.ToString() + "px";
                    }

                    foreach (var column in row.columns)
                    {
                        var td = new HtmlGenericControl("td");
                        tr.Controls.Add(td);
                        td.Attributes["class"] = column.@class;
                        td.Attributes["data-formula"] = column.formula;
                        td.Attributes["style"] = column.style;
                        td.InnerHtml = column.value;
                    }
                }
            }
        }

        public String ToString()
        {
            var builder = new StringBuilder();
            var writer = new StringWriter(builder);
            var html = new HtmlTextWriter(writer);

            RenderControl(html);
            return builder.ToString();
        }
    }
}
