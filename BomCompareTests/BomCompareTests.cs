using System.Diagnostics;
using System.Text;
using System.Xml;
using Aras.IOM;
using BomCompare = Hille.Aras.Core.BomCompare;

namespace BomCompareTests {
    public class BomCompareTests {

        private const string PART_ITEM_NUMBER = "P-1003";

        private Innovator innField;
        private Innovator Inn {
            get {
                if (innField is null) {
                    innField = InnovatorBase.getInnovator();
                }
                return innField;
            }
        }

        [Fact]
        public void BasicBomCompareTest() {
            // Get (latest part)
            Item part = GetPart(Inn, PART_ITEM_NUMBER);
            string currentId = part.getID();
            // Get its previous released revision
            Item previousReleasedRevision = GetPreviousReleasedItem(part);
            string idPreviousReleased = previousReleasedRevision.getID();

            // Setup an, index, i.e. what should be used as key for comparision
            BomCompare.IBomCompareItemProperty indexCompareProperty;
            indexCompareProperty = new BomCompare.RelationshipProperty("sort_order", "Sequence");

            // Create a new compare object
            var bomCompare = new BomCompare.PartBomCompare(Inn);
            // Add properties to compare
            List<BomCompare.IBomCompareItemProperty> compProperties = CompareProperties(indexCompareProperty);
            bomCompare.SetProperties(compProperties);

            // Set change description template
            bomCompare.ChangeDescription = " this value {0} compared value {1}";

            // Execute the compare
            bomCompare.Compare(currentId, idPreviousReleased, indexCompareProperty);

            // The compare now has populated BomCompareRows
            foreach (BomCompare.BomCompareRow row in bomCompare.BomCompareRows)
                Console.WriteLine(row.ChangeType.ToString());

            // Convert the BomCompareRow objects to XML
            System.Xml.XmlDocument xmlDoc;
            BomCompare.OutputFormat.IOutputSettings outputSettings = new BomCompare.OutputFormat.DefaultOutputSettings();
            outputSettings.CellColors.ChangedRowColor = "lightyellow";
            outputSettings.CellColors.ChangedCellColor = "yellow";
            xmlDoc = bomCompare.GetResultAsXml(outputSettings);

            // Convert xml to html and display
            string html = ConvertToHtml(xmlDoc);
            string tempFilePath = Path.Combine(Path.GetTempPath(), DateTime.Now.Ticks.ToString() + ".html");
            using (var sw = new StreamWriter(tempFilePath)) {
                sw.WriteLine(html);
            }

            // Open file to view and delete it afterwards
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(tempFilePath)
            {
                UseShellExecute = true
            };
            p.Start();
            System.Threading.Thread.Sleep(5000); // Wait five sec then delete file
            File.Delete(tempFilePath);

            Console.WriteLine("Done");

        }


        private Item GetPart(Innovator inn, string partNo) {
            string aml = "<AML><Item action='get' type='Part'><item_number>{0}</item_number></Item></AML>";
            aml = string.Format(aml, partNo);
            return inn.applyAML(aml);
        }

        private Item GetPreviousReleasedItem(Item releasedPart) {
            Item previousRev = default;
            string config_id = releasedPart.getProperty("config_id");
            // Ask for all generations within a revision
            Item allReleasedGenerations = Inn.newItem(releasedPart.getType(), "get");
            allReleasedGenerations.setAttribute("orderBy", "generation DESC");
            allReleasedGenerations.setProperty("config_id", config_id);
            allReleasedGenerations.setProperty("is_released", "1");
            allReleasedGenerations.setPropertyAttribute("id", "condition", "is not null");
            allReleasedGenerations = allReleasedGenerations.apply();
            int i = 0;
            while (i < allReleasedGenerations.getItemCount()) {
                Item generationItem = allReleasedGenerations.getItemByIndex(i);
                if (generationItem.getID != releasedPart.getID) {
                    // Previous generation
                    return generationItem;
                }
                i = i + 1;
            }

            return previousRev;
        }

        private string ConvertToHtml(XmlDocument xmlDoc) {
            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<meta charset='UTF-8'>");
            sb.AppendLine("<html>");

            sb.AppendLine(xmlDoc.GetElementsByTagName("table").Item(0).OuterXml);

            sb.AppendLine("</html>");
            return sb.ToString();
        }


        private List<BomCompare.IBomCompareItemProperty> CompareProperties(BomCompare.IBomCompareItemProperty indexProp) {
            var list = new List<BomCompare.IBomCompareItemProperty>();

            list.Add(indexProp);
            // Include but dont describe comparision on name
            var nameProp = new BomCompare.NonRelProperty("name", "Name");
            nameProp.IsComparable = false;

            list.Add(nameProp);
            list.Add(new BomCompare.RelationshipProperty("quantity", "Quantity"));
            list.Add(new BomCompare.NonRelProperty("item_number", "Number"));
            return list;
        }



    }
}