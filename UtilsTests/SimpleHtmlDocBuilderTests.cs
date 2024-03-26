﻿using me.fengyj.CommonLib.Utils;

namespace UtilsTests {
    [TestClass]
    public class SimpleHtmlDocBuilderTests {

        [TestMethod]
        public void Test() {

            var html = SimpleHtmlDocBuilder.Create()
                         .AppendHeader("hello ~")
                         .AppendParagraph("dfsdfsafasfsd")
                         .AppendTable(TableBuilder.Create(new string[] { "Col 1", "Col 2", "Col 3" })
                                     .AddRow("xyz", 1L, 3.4f)
                                     .AddRow(LinkBuilder.Create("http://bing.com", "bing"), 2000L, 523.001f))
                                .AppendHeader(HeaderBuilder.Create("color header: ").AppendInRed("red"))
                                .AppendParagraph(ParagraphBuilder.CreateInRed("warning: ")
                                .AppendInBlue("this is a test mail.")
                                .Append(LinkBuilder.Create("http://google.com", "google")))
                         .Build();

            Assert.IsNotNull(html);
        }
    }
}
