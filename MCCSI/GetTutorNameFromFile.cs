using System.IO;


namespace MCCSI
{
    class GetTutorNameFromFile
    {
        public string GetTutorNameFromFileMethod()
        {
            //read file to string variable

            

            string dataFile = File.ReadAllText(Form1.HiddenFile);
            string[] dataArray = dataFile.Split('|');


            string tutorsName = dataArray[0];

            return tutorsName;
            
        }
    }
}
