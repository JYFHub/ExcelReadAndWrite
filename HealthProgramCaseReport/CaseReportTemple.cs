using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HealthProgramCaseReport
{
    class CaseReportTemple
    {
        public string patientName { get; set; }
        public string patientAge { get; set; }
        public string patientBirthday { get; set; }
        public string patientIllnessNum { get; set; }
        public string patientConptionNum { get; set; }
        public string patientProductionNum { get; set; }
        public string patientLastMenstrualPeriod { get; set; }
        public string patientContraception { get; set; }
        public string patientCategory { get; set; }
        public string patientCheckAgainst { get; set; }
        public string patientTBS { get; set; }
        public string patientLeuCorrheaRegular { get; set; }
        public string patientHPV { get; set; }
        public List<string> patientVaginalImageList { get; set; }
        public List<string> patientLesionLocationList { get; set; }
        public List<string> patientEvaluatesImpressionsList { get; set; }
        public List<string> patientFurtherHandlingOfCommentsList { get; set; }

        public CaseReportTemple()
        {
            patientName = "";
            patientAge = "";
            patientBirthday = "";
            patientIllnessNum = "";
            patientConptionNum = "";
            patientProductionNum = "";
            patientLastMenstrualPeriod = "";
            patientContraception = "";
            patientCategory = "";
            patientCheckAgainst = "";
            patientTBS = "";
            patientLeuCorrheaRegular = "";
            patientHPV = "";
            if (patientVaginalImageList != null)
            {
                patientVaginalImageList.Clear();
            }
            else 
            { 
                patientVaginalImageList = new List<string>(); 
            }
            if (patientLesionLocationList != null)
            {
                patientLesionLocationList.Clear();
            }
            else
            {
                patientLesionLocationList = new List<string>();
            }
            if (patientEvaluatesImpressionsList != null)
            {
                patientEvaluatesImpressionsList.Clear();
            }
            else
            {
                patientEvaluatesImpressionsList = new List<string>();
            }
            if (patientFurtherHandlingOfCommentsList != null)
            {
                patientFurtherHandlingOfCommentsList.Clear();
            }
            else
            {
                patientFurtherHandlingOfCommentsList = new List<string>();
            }
        }
    }
}
