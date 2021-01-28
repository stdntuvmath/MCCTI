using System.IO;


namespace MCCSI
{
    class GetCheckBoxFlag
    {
        public string GetCheckBoxFlagMethod()
        {
            //read file to string variable



            string dataFile = File.ReadAllText(Form1.HiddenFile);
            string[] dataArray = dataFile.Split('|');


            string checkBoxState = dataArray[1];

            return checkBoxState;

        }
    }
}
