using System;

namespace OpenXmlPocDocker
{
    public class TestModelData
    {
        public static TestModelList GetList()
        {
            TestModelList tmList = new TestModelList
            {
                testData = new TestModel[100000]
            };

            for (int i = 0; i < 100000; i++)
            {
                tmList.testData[i] =
                 new TestModel
                 {
                     TestId = i,
                     TestName = $"Test{i}",
                     TestDesc = $"Tested {i} times",
                     TestDesc1 = i,
                     TestDesc2 = i,
                     TestDesc3 = i,
                     TestDesc4 = "TestDesc4",
                     TestDesc5 = "TestDesc5",
                     TestDesc6 = "TestDesc6",
                     TestDesc7 = "TestDesc7",
                     TestDesc8 = "TestDesc8",
                     TestDesc9 = "TestDesc9",
                     TestDesc10 = "TestDesc10",
                     TestDesc11 = "TestDesc11",
                     TestDesc12 = "TestDesc12",
                     TestDesc13 = "TestDesc13",
                     TestDesc14 = "TestDesc14",
                     TestDesc15 = "TestDesc15",
                     TestDesc16 = "TestDesc16",
                     TestDesc17 = "TestDesc17",
                     TestDesc18 = "TestDesc18",
                     TestDesc19 = "TestDesc19",
                     TestDesc20 = "TestDesc20",

                     TestDate = DateTime.Now.AddDays(-1)
                 };

            }

            return tmList;

        }
    }
}
