namespace ExcelMacroAdd.Models
{
    public class AdditionalDevices
    {
        public string vendor;
        public string shuntTrip24vArticle;
        public string shuntTrip48vArticle;
        public string shuntTrip230vArticle;
        public string undervoltageReleaseArticle;
        public string signalContactArticle;
        public string auxiliaryContactArticle;
        public string signalOrAuxiliaryContactArticle;

        public AdditionalDevices(string vendor, string shuntTrip24vArticle, string shuntTrip48vArticle, string shuntTrip230vArticle,
            string undervoltageReleaseArticle, string signalContactArticle, string auxiliaryContactArticle, string signalOrAuxiliaryContactArticle)
        {
            this.vendor = vendor;
            this.shuntTrip24vArticle = shuntTrip24vArticle;
            this.shuntTrip48vArticle = shuntTrip48vArticle;
            this.shuntTrip230vArticle = shuntTrip230vArticle;
            this.undervoltageReleaseArticle = undervoltageReleaseArticle;
            this.signalContactArticle = signalContactArticle;
            this.auxiliaryContactArticle = auxiliaryContactArticle;
            this.signalOrAuxiliaryContactArticle = signalOrAuxiliaryContactArticle;
        }

    }
}
