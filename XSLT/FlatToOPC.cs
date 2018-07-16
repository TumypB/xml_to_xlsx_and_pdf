using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Schema;

class OPC
{
    public static void FlatToOpc(XElement root, Stream output)
    {
        XNamespace pkg =
            "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace rel =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        using (Package package = Package.Open(output, FileMode.Create))
        {
            // add all parts (but not relationships)
            foreach (var xmlPart in root
                .Elements()
                .Where(p =>
                    (string)p.Attribute(pkg + "contentType") !=
                    "application/vnd.openxmlformats-package.relationships+xml"))
            {
                string name = (string)xmlPart.Attribute(pkg + "name");
                string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                if (contentType.EndsWith("xml"))
                {
                    Uri u = new Uri(name, UriKind.Relative);
                    PackagePart part = package.CreatePart(u, contentType,
                        CompressionOption.SuperFast);
                    using (Stream str = part.GetStream(FileMode.Create))
                    using (XmlWriter xmlWriter = XmlWriter.Create(str))
                        xmlPart.Element(pkg + "xmlData")
                            .Elements()
                            .First()
                            .WriteTo(xmlWriter);
                }
                else
                {
                    Uri u = new Uri(name, UriKind.Relative);
                    PackagePart part = package.CreatePart(u, contentType,
                        CompressionOption.SuperFast);
                    using (Stream str = part.GetStream(FileMode.Create))
                    using (BinaryWriter binaryWriter = new BinaryWriter(str))
                    {
                        string base64StringInChunks =
                            (string)xmlPart.Element(pkg + "binaryData");
                        char[] base64CharArray = base64StringInChunks
                            .Where(c => c != '\r' && c != '\n').ToArray();
                        byte[] byteArray =
                            System.Convert.FromBase64CharArray(base64CharArray,
                            0, base64CharArray.Length);
                        binaryWriter.Write(byteArray);
                    }
                }
            }

            foreach (var xmlPart in root.Elements())
            {
                string name = (string)xmlPart.Attribute(pkg + "name");
                string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                if (contentType ==
                    "application/vnd.openxmlformats-package.relationships+xml")
                {
                    // add the package level relationships
                    if (name == "/_rels/.rels")
                    {
                        foreach (XElement xmlRel in
                            xmlPart.Descendants(rel + "Relationship"))
                        {
                            string id = (string)xmlRel.Attribute("Id");
                            string type = (string)xmlRel.Attribute("Type");
                            string target = (string)xmlRel.Attribute("Target");
                            string targetMode =
                                (string)xmlRel.Attribute("TargetMode");
                            if (targetMode == "External")
                                package.CreateRelationship(
                                    new Uri(target, UriKind.Absolute),
                                    TargetMode.External, type, id);
                            else
                                package.CreateRelationship(
                                    new Uri(target, UriKind.Relative),
                                    TargetMode.Internal, type, id);
                        }
                    }
                    else
                    // add part level relationships
                    {
                        string directory = name.Substring(0, name.IndexOf("/_rels"));
                        string relsFilename = name.Substring(name.LastIndexOf('/'));
                        string filename =
                            relsFilename.Substring(0, relsFilename.IndexOf(".rels"));
                        PackagePart fromPart = package.GetPart(
                            new Uri(directory + filename, UriKind.Relative));
                        foreach (XElement xmlRel in
                            xmlPart.Descendants(rel + "Relationship"))
                        {
                            string id = (string)xmlRel.Attribute("Id");
                            string type = (string)xmlRel.Attribute("Type");
                            string target = (string)xmlRel.Attribute("Target");
                            string targetMode =
                                (string)xmlRel.Attribute("TargetMode");
                            if (targetMode == "External")
                                fromPart.CreateRelationship(
                                    new Uri(target, UriKind.Absolute),
                                    TargetMode.External, type, id);
                            else
                                fromPart.CreateRelationship(
                                    new Uri(target, UriKind.Relative),
                                    TargetMode.Internal, type, id);
                        }
                    }
                }
            }
        }
    }
}
