using System;
using PLCTools.Common;

static class Extensions
{
    /// <summary>
    /// Get substring of specified number of characters on the right.
    /// </summary>
    public static string Right(this string value, int length)
    {
        return value.Substring(value.Length - length);
    }
}
namespace PLCTools.Service
{
    class UserProcedure
    {
        private const int MaxStrands = 4;
        public void ExecuteProcedure(string ProcedureName, TagHandler Tags)
        {
            switch (ProcedureName.ToLower())
            {
                case "stageaccumulate": StageAccumulate(Tags); break;
                case "furnacedelays": FurnaceDelays(Tags); break;
                case "stagechanges": StageChanges(Tags); break;
                case "furnaceeoh": FurnaceEOH(Tags); break;
                case "furnacesoh": FurnaceSOH(Tags); break;
                case "billetcount": BilletCount(Tags); break;
                case "torchmansoh": TOR_SOH(Tags); break;
                case "torchmaneoh": TOR_EOH(Tags); break;
                case "furnacepwonoff": FurPowerOnOff(Tags); break;
                case "casterdelays": CasterDelays(Tags); break;
                case "accumulatestrandtimepwron_off": AccumulateStrandTimePwrOn_Off(Tags); break;
                case "calculatetimetundish": CalculateTimeTundish(Tags); break;
                case "callmillspplan": CallMillSPPlan(); break;
                case "whibmgapchange": WhiBMGapChange(Tags); break;
            }
        }
        private void StageAccumulate(TagHandler Tags)
        {
            int ActualStage;
            string ClearStages;

            int i;

            int[] Pwr_ON = new int[9];
            int[] Pwr_OFF = new int[9];
            float[] Elect_Kwh = new float[9];
            float[] Alt_Kwh = new float[9];
            float[] Lance1_O2 = new float[9];
            float[] Lance2_O2 = new float[9];
            float[] Burner1_O2 = new float[9];
            float[] Burner1_GAS = new float[9];
            float[] Burner1_SSO2 = new float[9];
            float[] Burner2_O2 = new float[9];
            float[] Burner2_GAS = new float[9];
            float[] Burner2_SSO2 = new float[9];
            float[] Burner3_O2 = new float[9];
            float[] Burner3_GAS = new float[9];
            float[] Burner3_SSO2 = new float[9];
            float[] Beak_O2 = new float[9];
            float[] Beak_Gas = new float[9];
            float[] Asbury_C = new float[9];
            float[] Asbury_East_C = new float[9];
            float[] Premier_C = new float[9];

            //float Elect_Kwh_Act;
            //float Alt_Kwh_Act;
            //float Lance1_O2_Act;
            //float Lance2_O2_Act;
            //float Burner1_O2_Act;
            //float Burner1_GAS_Act;
            //float Burner1_SSO2_Act;
            //float Burner2_O2_Act;
            //float Burner2_GAS_Act;
            //float Burner2_SSO2_Act;
            //float Burner3_O2_Act;
            //float Burner3_GAS_Act;
            //float Burner3_SSO2_Act;
            //float Beak_O2_Act;
            //float Beak_Gas_Act;
            //float Asbury_C_Act;
            //float Asbury_East_C_Act;
            //float Premier_C_Act;
            /*Parameter string in PLC2DB_Cfg
              FUR_ACTUAL_STAGE_NO FUR_AH_PWR_ON_MIN + FUR_AH_PWR_OFF_MIN + FUR_AH_ELECT_KWH + FUR_AH_LANCE1_O2_CON + FUR_AH_BURNER1_O2_CON + FUR_AH_BURNER1_GAS_CON + FUR_AH_BURNER1_SSO2_CON + FUR_AH_BURNER2_O2_CON + FUR_AH_BURNER2_GAS_CON + FUR_AH_BURNER2_SSO2_CON + FUR_AH_BURNER3_O2_CON + FUR_AH_BURNER3_GAS_CON + FUR_AH_BURNER3_SSO2_CON + FUR_AH_INJ1_CARBON + FUR_AH_INJ2_CARBON + FUR_AH_INJ3_CARBON
            */
            ActualStage = (int)Tags.GetTagsValue("FUR_ACTUAL_STAGE_NO");
            ClearStages = (string)Tags.GetTagsValue("INT_PRO_CLEAR_STAGES");
            //If ActualStage = 0 Then Stop
            if (ActualStage == 0 && ClearStages == "YES")
            {
                for (i = 0; i <= 8; i++)
                {
                    Pwr_ON[i] = 0;
                    Pwr_OFF[i] = 0;
                    Elect_Kwh[i] = 0;
                    Alt_Kwh[i] = 0;
                    Lance1_O2[i] = 0;
                    Lance2_O2[i] = 0;
                    Burner1_O2[i] = 0;
                    Burner1_GAS[i] = 0;
                    Burner1_SSO2[i] = 0;
                    Burner2_O2[i] = 0;
                    Burner2_GAS[i] = 0;
                    Burner2_SSO2[i] = 0;
                    Burner3_O2[i] = 0;
                    Burner3_GAS[i] = 0;
                    Burner3_SSO2[i] = 0;
                    Beak_O2[i] = 0;
                    Beak_Gas[i] = 0;
                    Asbury_C[i] = 0;
                    Asbury_East_C[i] = 0;
                    Premier_C[i] = 0;
                }
                Tags.UpdateTagsValue("INT_PRO_CLEAR_STAGES", "NO");
                UpdateTotalsStage(Tags, ref Pwr_ON, ref Pwr_OFF, ref Elect_Kwh, ref Alt_Kwh, ref Lance1_O2, ref Lance2_O2, ref Burner1_O2, ref Burner1_GAS, ref Burner1_SSO2, ref Burner2_O2, ref Burner2_GAS, ref Burner2_SSO2, ref Burner3_O2, ref Burner3_GAS, ref Burner3_SSO2, ref Beak_O2, ref Beak_Gas, ref Asbury_C, ref Asbury_East_C, ref Premier_C);
                Tags.UpdateTagsValue("INT_EV_PROC_STAGE_ACC_COMPLETED", 1);
            }
        }
        private void RetrieveStageValues(TagHandler Tags, ref int[] Pwr_ON, ref int[] Pwr_OFF, ref float[] Elect_Kwh, ref float[] Alt_Kwh, ref float[] Lance1_O2, ref float[] Lance2_O2, ref float[] Burner1_O2, ref float[] Burner1_GAS, ref float[] Burner1_SSO2, ref float[] Burner2_O2, ref float[] Burner2_GAS, ref float[] Burner2_SSO2, ref float[] Burner3_O2, ref float[] Burner3_GAS, ref float[] Burner3_SSO2, ref float[] Beak_O2, ref float[] Beak_Gas, ref float[] Asbury_C, ref float[] Asbury_East_C, ref float[] Premier_C)
        {
            Pwr_ON[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG00");
            Pwr_ON[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG10");
            Pwr_ON[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG12");
            Pwr_ON[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG20");
            Pwr_ON[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG22");
            Pwr_ON[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG30");
            Pwr_ON[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG40");
            Pwr_ON[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG50");
            Pwr_ON[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_PWR_ON_STG80");

            Pwr_OFF[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG00");
            Pwr_OFF[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG10");
            Pwr_OFF[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG12");
            Pwr_OFF[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG20");
            Pwr_OFF[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG22");
            Pwr_OFF[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG30");
            Pwr_OFF[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG40");
            Pwr_OFF[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG50");
            Pwr_OFF[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_PWR_OFF_STG80");

            Elect_Kwh[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG00");
            Elect_Kwh[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG10");
            Elect_Kwh[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG12");
            Elect_Kwh[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG20");
            Elect_Kwh[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG22");
            Elect_Kwh[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG30");
            Elect_Kwh[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG40");
            Elect_Kwh[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG50");
            Elect_Kwh[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_ELECT_KWH_STG80");

            Alt_Kwh[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG00");
            Alt_Kwh[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG10");
            Alt_Kwh[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG12");
            Alt_Kwh[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG20");
            Alt_Kwh[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG22");
            Alt_Kwh[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG30");
            Alt_Kwh[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG40");
            Alt_Kwh[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG50");
            Alt_Kwh[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_ALT_KWH_STG80");

            Lance1_O2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG00");
            Lance1_O2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG10");
            Lance1_O2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG12");
            Lance1_O2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG20");
            Lance1_O2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG22");
            Lance1_O2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG30");
            Lance1_O2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG40");
            Lance1_O2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG50");
            Lance1_O2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_LANCE1_O2_STG80");

            Lance2_O2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG00");
            Lance2_O2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG10");
            Lance2_O2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG12");
            Lance2_O2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG20");
            Lance2_O2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG22");
            Lance2_O2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG30");
            Lance2_O2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG40");
            Lance2_O2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG50");
            Lance2_O2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_LANCE2_O2_STG80");

            Burner1_O2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG00");
            Burner1_O2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG10");
            Burner1_O2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG12");
            Burner1_O2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG20");
            Burner1_O2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG22");
            Burner1_O2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG30");
            Burner1_O2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG40");
            Burner1_O2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG50");
            Burner1_O2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_O2_STG80");

            Burner1_GAS[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG00");
            Burner1_GAS[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG10");
            Burner1_GAS[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG12");
            Burner1_GAS[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG20");
            Burner1_GAS[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG22");
            Burner1_GAS[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG30");
            Burner1_GAS[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG40");
            Burner1_GAS[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG50");
            Burner1_GAS[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_GAS_STG80");

            Burner1_SSO2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG00");
            Burner1_SSO2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG10");
            Burner1_SSO2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG12");
            Burner1_SSO2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG20");
            Burner1_SSO2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG22");
            Burner1_SSO2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG30");
            Burner1_SSO2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG40");
            Burner1_SSO2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG50");
            Burner1_SSO2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER1_SSO2_STG80");

            Burner2_O2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG00");
            Burner2_O2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG10");
            Burner2_O2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG12");
            Burner2_O2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG20");
            Burner2_O2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG22");
            Burner2_O2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG30");
            Burner2_O2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG40");
            Burner2_O2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG50");
            Burner2_O2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_O2_STG80");

            Burner2_GAS[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG00");
            Burner2_GAS[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG10");
            Burner2_GAS[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG12");
            Burner2_GAS[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG20");
            Burner2_GAS[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG22");
            Burner2_GAS[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG30");
            Burner2_GAS[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG40");
            Burner2_GAS[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG50");
            Burner2_GAS[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_GAS_STG80");

            Burner2_SSO2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG00");
            Burner2_SSO2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG10");
            Burner2_SSO2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG12");
            Burner2_SSO2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG20");
            Burner2_SSO2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG22");
            Burner2_SSO2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG30");
            Burner2_SSO2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG40");
            Burner2_SSO2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG50");
            Burner2_SSO2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER2_SSO2_STG80");

            Burner3_O2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG00");
            Burner3_O2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG10");
            Burner3_O2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG12");
            Burner3_O2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG20");
            Burner3_O2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG22");
            Burner3_O2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG30");
            Burner3_O2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG40");
            Burner3_O2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG50");
            Burner3_O2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_O2_STG80");

            Burner3_GAS[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG00");
            Burner3_GAS[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG10");
            Burner3_GAS[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG12");
            Burner3_GAS[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG20");
            Burner3_GAS[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG22");
            Burner3_GAS[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG30");
            Burner3_GAS[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG40");
            Burner3_GAS[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG50");
            Burner3_GAS[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_GAS_STG80");

            Burner3_SSO2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG00");
            Burner3_SSO2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG10");
            Burner3_SSO2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG12");
            Burner3_SSO2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG20");
            Burner3_SSO2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG22");
            Burner3_SSO2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG30");
            Burner3_SSO2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG40");
            Burner3_SSO2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG50");
            Burner3_SSO2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BURNER3_SSO2_STG80");

            Beak_O2[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG00");
            Beak_O2[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG10");
            Beak_O2[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG12");
            Beak_O2[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG20");
            Beak_O2[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG22");
            Beak_O2[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG30");
            Beak_O2[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG40");
            Beak_O2[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG50");
            Beak_O2[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_O2_STG80");

            Beak_Gas[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG00");
            Beak_Gas[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG10");
            Beak_Gas[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG10");
            Beak_Gas[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG20");
            Beak_Gas[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG20");
            Beak_Gas[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG30");
            Beak_Gas[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG40");
            Beak_Gas[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG50");
            Beak_Gas[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_BEAK_GAS_STG80");

            Asbury_C[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG00");
            Asbury_C[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG10");
            Asbury_C[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG12");
            Asbury_C[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG20");
            Asbury_C[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG22");
            Asbury_C[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG30");
            Asbury_C[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG40");
            Asbury_C[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG50");
            Asbury_C[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_C_STG80");

            Asbury_East_C[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG00");
            Asbury_East_C[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG10");
            Asbury_East_C[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG12");
            Asbury_East_C[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG20");
            Asbury_East_C[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG22");
            Asbury_East_C[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG30");
            Asbury_East_C[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG40");
            Asbury_East_C[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG50");
            Asbury_East_C[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_ASBURY_EAST_C_STG80");

            Premier_C[StageIndex(0)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG00");
            Premier_C[StageIndex(10)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG10");
            Premier_C[StageIndex(12)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG12");
            Premier_C[StageIndex(20)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG20");
            Premier_C[StageIndex(22)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG22");
            Premier_C[StageIndex(30)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG30");
            Premier_C[StageIndex(40)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG40");
            Premier_C[StageIndex(50)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG50");
            Premier_C[StageIndex(80)] = (int)Tags.GetTagsValue("INT_FCE_PREMIER_C_STG80");
        }
        private void UpdateTotalsStage(TagHandler Tags, ref int[] Pwr_ON, ref int[] Pwr_OFF, ref float[] Elect_Kwh, ref float[] Alt_Kwh, ref float[] Lance1_O2, ref float[] Lance2_O2, ref float[] Burner1_O2, ref float[] Burner1_GAS, ref float[] Burner1_SSO2, ref float[] Burner2_O2, ref float[] Burner2_GAS, ref float[] Burner2_SSO2, ref float[] Burner3_O2, ref float[] Burner3_GAS, ref float[] Burner3_SSO2, ref float[] Beak_O2, ref float[] Beak_Gas, ref float[] Asbury_C, ref float[] Asbury_East_C, ref float[] Premier_C)
        {
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG00", Pwr_ON[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG10", Pwr_ON[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG12", Pwr_ON[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG20", Pwr_ON[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG22", Pwr_ON[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG30", Pwr_ON[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG40", Pwr_ON[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG50", Pwr_ON[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_ON_STG80", Pwr_ON[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG00", Pwr_OFF[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG10", Pwr_OFF[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG12", Pwr_OFF[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG20", Pwr_OFF[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG22", Pwr_OFF[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG30", Pwr_OFF[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG40", Pwr_OFF[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG50", Pwr_OFF[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_PWR_OFF_STG80", Pwr_OFF[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG00", Elect_Kwh[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG10", Elect_Kwh[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG12", Elect_Kwh[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG20", Elect_Kwh[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG22", Elect_Kwh[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG30", Elect_Kwh[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG40", Elect_Kwh[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG50", Elect_Kwh[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_ELECT_KWH_STG80", Elect_Kwh[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG00", Alt_Kwh[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG10", Alt_Kwh[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG12", Alt_Kwh[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG20", Alt_Kwh[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG22", Alt_Kwh[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG30", Alt_Kwh[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG40", Alt_Kwh[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG50", Alt_Kwh[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_ALT_KWH_STG80", Alt_Kwh[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG00", Lance1_O2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG10", Lance1_O2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG12", Lance1_O2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG20", Lance1_O2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG22", Lance1_O2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG30", Lance1_O2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG40", Lance1_O2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG50", Lance1_O2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE1_O2_STG80", Lance1_O2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG00", Lance2_O2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG10", Lance2_O2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG12", Lance2_O2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG20", Lance2_O2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG22", Lance2_O2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG30", Lance2_O2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG40", Lance2_O2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG50", Lance2_O2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_LANCE2_O2_STG80", Lance2_O2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG00", Burner1_O2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG10", Burner1_O2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG12", Burner1_O2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG20", Burner1_O2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG22", Burner1_O2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG30", Burner1_O2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG40", Burner1_O2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG50", Burner1_O2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_O2_STG80", Burner1_O2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG00", Burner1_GAS[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG10", Burner1_GAS[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG12", Burner1_GAS[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG20", Burner1_GAS[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG22", Burner1_GAS[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG30", Burner1_GAS[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG40", Burner1_GAS[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG50", Burner1_GAS[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_GAS_STG80", Burner1_GAS[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG00", Burner1_SSO2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG10", Burner1_SSO2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG12", Burner1_SSO2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG20", Burner1_SSO2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG22", Burner1_SSO2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG30", Burner1_SSO2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG40", Burner1_SSO2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG50", Burner1_SSO2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER1_SSO2_STG80", Burner1_SSO2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG00", Burner2_O2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG10", Burner2_O2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG12", Burner2_O2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG20", Burner2_O2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG22", Burner2_O2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG30", Burner2_O2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG40", Burner2_O2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG50", Burner2_O2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_O2_STG80", Burner2_O2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG00", Burner2_GAS[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG10", Burner2_GAS[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG12", Burner2_GAS[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG20", Burner2_GAS[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG22", Burner2_GAS[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG30", Burner2_GAS[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG40", Burner2_GAS[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG50", Burner2_GAS[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_GAS_STG80", Burner2_GAS[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG00", Burner2_SSO2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG10", Burner2_SSO2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG12", Burner2_SSO2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG20", Burner2_SSO2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG22", Burner2_SSO2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG30", Burner2_SSO2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG40", Burner2_SSO2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG50", Burner2_SSO2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER2_SSO2_STG80", Burner2_SSO2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG00", Burner3_O2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG10", Burner3_O2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG12", Burner3_O2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG20", Burner3_O2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG22", Burner3_O2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG30", Burner3_O2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG40", Burner3_O2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG50", Burner3_O2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_O2_STG80", Burner3_O2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG00", Burner3_GAS[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG10", Burner3_GAS[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG12", Burner3_GAS[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG20", Burner3_GAS[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG22", Burner3_GAS[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG30", Burner3_GAS[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG40", Burner3_GAS[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG50", Burner3_GAS[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_GAS_STG80", Burner3_GAS[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG00", Burner3_SSO2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG10", Burner3_SSO2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG12", Burner3_SSO2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG20", Burner3_SSO2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG22", Burner3_SSO2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG30", Burner3_SSO2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG40", Burner3_SSO2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG50", Burner3_SSO2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BURNER3_SSO2_STG80", Burner3_SSO2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG00", Beak_O2[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG10", Beak_O2[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG12", Beak_O2[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG20", Beak_O2[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG22", Beak_O2[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG30", Beak_O2[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG40", Beak_O2[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG50", Beak_O2[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_O2_STG80", Beak_O2[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG00", Beak_Gas[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG10", Beak_Gas[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG12", Beak_Gas[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG20", Beak_Gas[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG22", Beak_Gas[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG30", Beak_Gas[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG40", Beak_Gas[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG50", Beak_Gas[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_BEAK_GAS_STG80", Beak_Gas[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG00", Asbury_C[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG10", Asbury_C[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG12", Asbury_C[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG20", Asbury_C[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG22", Asbury_C[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG30", Asbury_C[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG40", Asbury_C[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG50", Asbury_C[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_C_STG80", Asbury_C[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG00", Asbury_East_C[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG10", Asbury_East_C[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG12", Asbury_East_C[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG20", Asbury_East_C[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG22", Asbury_East_C[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG30", Asbury_East_C[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG40", Asbury_East_C[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG50", Asbury_East_C[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_ASBURY_EAST_C_STG80", Asbury_East_C[StageIndex(80)]);

            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG00", Premier_C[StageIndex(0)]);
            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG10", Premier_C[StageIndex(10)]);
            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG12", Premier_C[StageIndex(12)]);
            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG20", Premier_C[StageIndex(20)]);
            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG22", Premier_C[StageIndex(22)]);
            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG30", Premier_C[StageIndex(30)]);
            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG40", Premier_C[StageIndex(40)]);
            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG50", Premier_C[StageIndex(50)]);
            Tags.UpdateTagsValue("INT_FCE_PREMIER_C_STG80", Premier_C[StageIndex(80)]);
        }
        private void FurnaceDelays(TagHandler Tags)
        {
            int ActualStage;
            string StageDesc = "";
            int PrevStage;
            int PowerOn;
            DateTime HeatStartTime;
            DateTime StartTime;
            DateTime EndTime;
            int TimeDelay;
            int DurDelay;
            int DelayStatus;

            int MinutesStartDelay;
            int SecondsStartDelay;
            int MinutesEndDelay;
            int SecondsEndDelay;
            int MinutesDurDelay;
            int SecondsDurDelay;


            
            PowerOn = (int)Tags.GetTagsValue("FUR_POWER_ON");
            ActualStage = (int)Tags.GetTagsValue("FUR_ACTUAL_STAGE_NO");
            PrevStage = (int)Tags.GetTagsValue("INT_FCE_PREV_STAGE_DELAY");
            HeatStartTime = (DateTime)Tags.GetTagsValue("INT_FCE_HEAT_TIME");
            DelayStatus = (int)Tags.GetTagsValue("INT_FCE_DELAY_STATUS");

            string DelayCode;// = "2099"; // No reason entered


            if (PowerOn == 0)
            {
                switch (ActualStage)
                {
                    case 0 // Turnaround
                   :
                        {
                            StageDesc = "TA";
                            break;
                        }

                    case 10 // 1st Charge
                    :
                        {
                            StageDesc = "CH1";
                            break;
                        }

                    case 12 // 1st Back Charge
             :
                        {
                            StageDesc = "BC1";
                            break;
                        }

                    case 20 // 2nd Charge
             :
                        {
                            StageDesc = "CH2";
                            break;
                        }

                    case 22 // 2nd Back Charge
             :
                        {
                            StageDesc = "BC2";
                            break;
                        }

                    case 30 // 3rd Charge
             :
                        {
                            StageDesc = "CH3";
                            break;
                        }

                    case 40 // Refining
             :
                        {
                            StageDesc = "REF";
                            break;
                        }

                    case 50 // Tapping
             :
                        {
                            StageDesc = "TAP";
                            break;
                        }
                }

                // -- Solves the program startup
                if (HeatStartTime == null)
                {
                    HeatStartTime = DateTime.Now;
                    Tags.UpdateTagsValue("INT_FCE_HEAT_TIME", HeatStartTime);
                }

                StartTime = DateTime.Now;
                TimeDelay = (int)(HeatStartTime - StartTime).TotalSeconds;
                MinutesStartDelay = TimeDelay / 60;
                SecondsStartDelay = TimeDelay % 60;

                PrevStage = ActualStage;

                DelayStatus = 1;

                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME", StartTime);
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_MIN", MinutesStartDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_SEC", SecondsStartDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_STATUS", DelayStatus);

                Tags.UpdateTagsValue("INT_FCE_STAGE_DESC", StageDesc);
                Tags.UpdateTagsValue("INT_FCE_PREV_STAGE_DELAY", PrevStage);

                Tags.UpdateTagsValue("INT_EV_PROC_FCE_DELAY_INSERT", 1);
            }
            else if (DelayStatus == 1 & HeatStartTime != null)
            {
                StartTime = (DateTime)Tags.GetTagsValue("INT_FCE_DELAY_START_TIME");
                EndTime = DateTime.Now;

                TimeDelay = (int)(HeatStartTime - EndTime).TotalSeconds;
                MinutesEndDelay = TimeDelay / 60;
                SecondsEndDelay = TimeDelay % 60;

                DurDelay = (int)(StartTime - EndTime).TotalSeconds;

                MinutesDurDelay = DurDelay / 60;
                SecondsDurDelay = DurDelay % 60;

                DelayCode = "2099";

                switch (PrevStage)
                {
                    case 0 // Turnaround
                   :
                        {
                            DelayCode = "126";
                            Tags.UpdateTagsValue("INT_FCE_TURNAROUND_DUR_SEC", DurDelay);
                            break;
                        }
                }



                DelayStatus = 0;

                Tags.UpdateTagsValue("INT_FCE_DELAY_END_TIME", EndTime);
                Tags.UpdateTagsValue("INT_FCE_DELAY_END_TIME_MIN", MinutesEndDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_END_TIME_SEC", SecondsEndDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_DURATION_MIN", MinutesDurDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_DURATION_SEC", SecondsDurDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_NO", DelayCode);
                Tags.UpdateTagsValue("INT_EV_PROC_FCE_DELAY_UPDATE", 1);
                Tags.UpdateTagsValue("INT_FCE_DELAY_STATUS", DelayStatus);
            }
        }
        private void StageChanges(TagHandler Tags)
        {
            // Raised by INT_EV_FCE_STAGE_CHANGED = (FUR_ACTUAL_STAGE_NO <>INT_FCE_PREV_STAGE) * (1) on the CFG

            int ActualStage;
            int PrevStage;
            int PrevStageDelay;
            int PowerOn;
            DateTime HeatStartTime;
            DateTime StartTime;
            DateTime PrevStartTime;

            DateTime EndTime;
            int TimeDelay;
            int DurDelay;

            int MinutesStartDelay;
            int SecondsStartDelay;
            int MinutesPrevStartDelay;
            int SecondsPrevStartDelay;

            //int MinutesEndDelay;
            //int SecondsEndDelay;
            int MinutesDurDelay;
            int SecondsDurDelay;

            int ActualHeat_1;
            int ActualHeat_2;
            int PrevHeat_1;
            int PrevHeat_2;

            PowerOn = (int)Tags.GetTagsValue("FUR_POWER_ON");
            ActualStage = (int)Tags.GetTagsValue("FUR_ACTUAL_STAGE_NO");
            //PrevStage = (int)Tags.GetTagsValue("INT_FCE_PREV_STAGE");
            PrevStageDelay = (int)Tags.GetTagsValue("INT_FCE_PREV_STAGE_DELAY");

            // --- Setea el valor del Stage Anterior para el pr�ximo ciclo.
            PrevStage = ActualStage;

            if (ActualStage == 0)
                // Must perform EOH y SOH

                // Raises the event EOH
                // The EOH procedure raises SOH when complete
                Tags.UpdateTagsValue("INT_EV_FCE_EOH", 1);

            if (ActualStage == 50)
            {
                // Preserves the Actual Heat as Previous Heat to update Delays
                ActualHeat_1 = (int)Tags.GetTagsValue("FUR_AC_HEAT_YEAR");
                ActualHeat_2 = (int)Tags.GetTagsValue("FUR_HEAT_ACTUAL");

                PrevHeat_1 = ActualHeat_1;
                PrevHeat_2 = ActualHeat_2;

                Tags.UpdateTagsValue("INT_FCE_PREV_HEAT_1", PrevHeat_1);
                Tags.UpdateTagsValue("INT_FCE_PREV_HEAT_2", PrevHeat_2);
            }

            // Checks if there is a furnace delay started on Tap and Turnaround
            if (PrevStageDelay == 50 | PrevStageDelay == 40 & ActualStage == 0 & PowerOn == 0)
            {

                // Preparara TAGS para cerrar delay anterior.
                StartTime = (DateTime)Tags.GetTagsValue("INT_FCE_DELAY_START_TIME");
                MinutesStartDelay = (int)Tags.GetTagsValue("INT_FCE_DELAY_START_TIME_MIN");
                SecondsStartDelay = (int)Tags.GetTagsValue("INT_FCE_DELAY_START_TIME_SEC");
                HeatStartTime = (DateTime)Tags.GetTagsValue("INT_FCE_HEAT_TIME");

                PrevStartTime = StartTime;
                MinutesPrevStartDelay = MinutesStartDelay;
                SecondsPrevStartDelay = SecondsStartDelay;

                EndTime = DateTime.Now;

                //TimeDelay = (int)(HeatStartTime - EndTime).TotalSeconds;
                //MinutesEndDelay = TimeDelay / 60;
                //SecondsEndDelay = TimeDelay % 60;

                DurDelay = (int)(StartTime - EndTime).TotalSeconds;
                MinutesDurDelay = DurDelay / 60;
                SecondsDurDelay = DurDelay % 60;


                // Preparara TAGS para insertar un nuevo delay
                // Dispara el Delay del Tap ahora, y deja listo el turnaround para hacerlo con el SOH

                StartTime = DateTime.Now;
                TimeDelay = (int)(HeatStartTime - StartTime).TotalSeconds;
                MinutesStartDelay = TimeDelay / 60;
                SecondsStartDelay = TimeDelay % 60;

                PrevStage = ActualStage;

                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_PREV", PrevStartTime);
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_MIN_PREV", MinutesPrevStartDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_SEC_PREV", SecondsPrevStartDelay);

                Tags.UpdateTagsValue("INT_FCE_DELAY_DURATION_MIN", MinutesDurDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_DURATION_SEC", SecondsDurDelay);

                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME", StartTime);
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_MIN", MinutesStartDelay);
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_SEC", SecondsStartDelay);


                Tags.UpdateTagsValue("INT_FCE_PREV_STAGE_DELAY", PrevStage);

                Tags.UpdateTagsValue("INT_EV_PROC_FCE_DELAY_TAP", 1);
                // Call Global.setValueRtData(Global.tagsRTDATA, "INT_EV_PROC_FCE_DELAY_TURNAROUND", 1)
                Tags.UpdateTagsValue("INT_FCE_DELAY_STATUS", 1);
            }

            Tags.UpdateTagsValue("INT_FCE_PREV_STAGE", PrevStage);
        }
        private void FurnaceEOH(TagHandler Tags)
        {
            // End Of Heat Procedure
            // When finish, raises the Start of Heat
            Tags.UpdateTagsValue("INT_EV_FCE_SOH", 1);
        }
        private void FurnaceSOH(TagHandler Tags)
        {
            // Start Of Heat Procedure
            DateTime HeatStartTime;
            DateTime LastTimeUpdate;
            int PowerOn;

            HeatStartTime = DateTime.Now;
            LastTimeUpdate = DateTime.Now;

            Tags.UpdateTagsValue("INT_FCE_HEAT_TIME", HeatStartTime);
            Tags.UpdateTagsValue("INT_FUR_LAST_POWER_TIME", LastTimeUpdate);

            Tags.UpdateTagsValue("INT_FCE_POWER_ON_SEC", 0);
            Tags.UpdateTagsValue("INT_FCE_POWER_OFF_SEC", 0);

            // Dispara el Delay del Turnaround
            PowerOn = (int)Tags.GetTagsValue("FUR_POWER_ON");
            if (PowerOn == 0)
            {
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME", DateTime.Now);
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_MIN", 0);
                Tags.UpdateTagsValue("INT_FCE_DELAY_START_TIME_SEC", 0);
                Tags.UpdateTagsValue("INT_FCE_DELAY_STATUS", 1);

                Tags.UpdateTagsValue("INT_EV_PROC_FCE_DELAY_TURNAROUND", 1);
            }
            Tags.UpdateTagsValue("INT_FCE_TURNAROUND_DUR_SEC", 0);
        }
        private void FurPowerOnOff(TagHandler Tags)
        {
            AccumulateTimePwrOn_Off(true, Tags);
        }
        private void AccumulateTimePwrOn_Off(bool ChangeStatus, TagHandler Tags)
        {
            // ChangeStatus means: if is calling from the event of change the Power On/Off of the furnace
            // Or just from other procedures to accumulate times.
            // If is from the ChangeStatus, must accumulate to the previous status.

            DateTime LastTimeUpdate;
            int PowerOn;
            int SecDiff;
            int PowerOnSec;
            int PowerOffSec;

            PowerOn = (int)Tags.GetTagsValue("FUR_POWER_ON");
            LastTimeUpdate = (DateTime)Tags.GetTagsValue("INT_FCE_LAST_POWER_TIME");
            PowerOnSec = (int)Tags.GetTagsValue("INT_FCE_POWER_ON_SEC");
            PowerOffSec = (int)Tags.GetTagsValue("INT_FCE_POWER_OFF_SEC");

            if (LastTimeUpdate == null)
                LastTimeUpdate = DateTime.Now;

            SecDiff = (int)(LastTimeUpdate - DateTime.Now).TotalSeconds;

            if ((PowerOn == 1 & !(ChangeStatus)) | (PowerOn == 0 & ChangeStatus))
                PowerOnSec += SecDiff;
            else
                PowerOffSec += SecDiff;

            LastTimeUpdate = DateTime.Now;

            Tags.UpdateTagsValue("INT_FCE_LAST_POWER_TIME", LastTimeUpdate);
            Tags.UpdateTagsValue("INT_FCE_POWER_ON_SEC", PowerOnSec);
            Tags.UpdateTagsValue("INT_FCE_POWER_OFF_SEC", PowerOffSec);
        }
        private void CasterDelays(TagHandler Tags)
        {

            // Const MaxStrands = 4

            int StrandDelay;
            int LadleOpen;
            int i;

            StrandDelay = (int)Tags.GetTagsValue("INT_CAS_DELAY_STRAND");

            LadleOpen = (int)Tags.GetTagsValue("CCM_EV_LDL_OPENED");

            if (StrandDelay > MaxStrands)
            {
                for (i = 1; i <= MaxStrands; i++)
                    CasterStartEndDelay(i, LadleOpen, Tags);
            }
            else if (StrandDelay > 0)
                CasterStartEndDelay(StrandDelay, LadleOpen, Tags);
        }
        private void CasterStartEndDelay(int StrandDelay, int LadleOpen, TagHandler Tags)
        {
            int isCasting;
            int DelayStatus;

            DateTime HeatStartTime;
            DateTime PrevHeatStartTime;
            DateTime TimeStartDelay;
            int TimeInHeatDelayStart;
            int TimeInHeatDelayStartMin;
            int TimeInHeatDelayStartSec;
            int DelayDuration;
            int TimeInHeatDelayEnd;
            int TotalHeatDelayDuration;

            string Tag1;

            HeatStartTime = (DateTime)Tags.GetTagsValue("INT_CAS_HEAT_TIME");
            // Solver program start
            if (HeatStartTime == null)
                HeatStartTime = DateTime.Now;

            PrevHeatStartTime = (DateTime)Tags.GetTagsValue("INT_CAS_HEAT_TIME_PREV");
            if (PrevHeatStartTime == null)
                PrevHeatStartTime = DateTime.Now;

            Tag1 = "TOR_STR" + StrandDelay + "_OPEN";
            isCasting = (int)Tags.GetTagsValue(Tag1);

            Tag1 = "INT_CAS_DELAY_STATUS_STR" + StrandDelay.ToString().Trim();
            DelayStatus = (int)Tags.GetTagsValue(Tag1);

            // ---- Casos para cerar o abrir Delays
            // Si LadleOpen = 1 y DelayStatus = 2, esta iniciando una colada, con delays que empezaron
            // despues de que cerro la colada anterior.
            if (DelayStatus == 2 & LadleOpen == 1)
            {
                Tag1 = "INT_CAS_DELAY_TIME_START_STR" + StrandDelay.ToString().Trim();
                TimeStartDelay = (DateTime)Tags.GetTagsValue(Tag1);
                if (TimeStartDelay == null)
                    DelayDuration = 0;
                else
                    DelayDuration = (int)(TimeStartDelay - DateTime.Now).TotalSeconds;

                Tag1 = "INT_CAS_DELAY_DUR_TOT_STR" + StrandDelay.ToString().Trim();
                TotalHeatDelayDuration = (int)Tags.GetTagsValue(Tag1);
                TotalHeatDelayDuration += DelayDuration;

                Tag1 = "INT_CAS_DELAY_DUR_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, DelayDuration);

                TimeInHeatDelayEnd = (int)(PrevHeatStartTime - DateTime.Now).TotalSeconds;
                Tag1 = "INT_CAS_DELAY_TIH_END_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, TimeInHeatDelayEnd);

                Tag1 = "INT_CAS_DELAY_DUR_TOT_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, TotalHeatDelayDuration);

                Tag1 = "INT_CAS_DELAY_STATUS_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, 0);

                Tag1 = "INT_EV_CAS_DELAY_UPDATE_PR_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, 1);
            }

            // If (isCasting = 1 Or (LadleOpen = 0 And isCasting = 0)) And DelayStatus = 1 Then ' Ends Delay
            if (DelayStatus == 1)
            {
                Tag1 = "INT_CAS_DELAY_TIME_START_STR" + StrandDelay.ToString().Trim();
                TimeStartDelay = (DateTime)Tags.GetTagsValue(Tag1);
                if (TimeStartDelay == null)
                    DelayDuration = 0;
                else
                    DelayDuration = (int)(TimeStartDelay - DateTime.Now).TotalSeconds;

                Tag1 = "INT_CAS_DELAY_DUR_TOT_STR" + StrandDelay.ToString().Trim();
                TotalHeatDelayDuration = (int)Tags.GetTagsValue(Tag1);
                TotalHeatDelayDuration += DelayDuration;

                Tag1 = "INT_CAS_DELAY_DUR_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, DelayDuration);

                TimeInHeatDelayEnd = (int)(HeatStartTime - DateTime.Now).TotalSeconds;
                Tag1 = "INT_CAS_DELAY_TIH_END_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, TimeInHeatDelayEnd);

                Tag1 = "INT_CAS_DELAY_DUR_TOT_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, TotalHeatDelayDuration);

                Tag1 = "INT_CAS_DELAY_STATUS_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, 0);

                Tag1 = "INT_EV_CAS_DELAY_UPDATE_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, 1);
            }
            else if (isCasting == 0)
            {
                TimeStartDelay = DateTime.Now;
                TimeInHeatDelayStart = (int)(HeatStartTime - TimeStartDelay).TotalSeconds;
                TimeInHeatDelayStartMin = TimeInHeatDelayStart / 60;
                TimeInHeatDelayStartSec = TimeInHeatDelayStart % 60;

                Tag1 = "INT_CAS_DELAY_TIME_START_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, TimeStartDelay);

                Tag1 = "INT_CAS_DELAY_TIH_START_MIN_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, TimeInHeatDelayStartMin);

                Tag1 = "INT_CAS_DELAY_TIH_START_SEC_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, TimeInHeatDelayStartSec);

                if (LadleOpen == 1)
                    DelayStatus = 1;
                else
                    DelayStatus = 2;
                Tag1 = "INT_CAS_DELAY_STATUS_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, DelayStatus);

                Tag1 = "INT_EV_CAS_DELAY_INSERT_STR" + StrandDelay.ToString().Trim();
                Tags.UpdateTagsValue(Tag1, 1);
            }
        }
        private void BilletCount(TagHandler Tags)
        {
            // Check if the counter changes per each strand and triggs the correspondant uptades
            // Check Also if the new billets beints to a different Heat and triggs the EOH and SOH on Torchman

            // Const MaxStrands = 4

            int[] ActualCount = new int[MaxStrands + 1];
            int[] PrevCount = new int[MaxStrands + 1];
            int[] AcHeatYear = new int[MaxStrands + 1];
            int[] AcHeatNumber = new int[MaxStrands + 1];

            string Tag1;
            string Tag2;




            string Evnt;
            bool HeatChange;
            bool isCasting;

            int TorActHeatYear = (int)Tags.GetTagsValue("INT_TOR_ACTUAL_HEAT_1");
            int TorActHeatNumber = (int)Tags.GetTagsValue("INT_TOR_ACTUAL_HEAT_2");
            //int TorPrHeatYear = (int)Tags.GetTagsValue("INT_TOR_PREV_HEAT_1");
            //int TorPrHeatNumber = (int)Tags.GetTagsValue("INT_TOR_PREV_HEAT_2");

            int TorAcHeat = TorActHeatYear * 10000 + TorActHeatNumber;
            //int TorPrHeat = TorPrHeatYear * 10000 + TorPrHeatNumber;

            // Solves the program Startup
            //if (TorPrHeat == 0) TorPrHeat = TorAcHeat;

            int i;
            HeatChange = false;
            for (i = 1; i <= MaxStrands; i++)
            {
                // ---------- Check the Counter
                Tag1 = "TOR_STR" + Convert.ToString(i).Trim() + "_BILLET_COUNT";
                ActualCount[i] = (int)Tags.GetTagsValue(Tag1);

                Tag2 = "INT_TOR_STR" + Convert.ToString(i).Trim() + "_LAST_BILLET_COUNT";
                PrevCount[i] = (int)Tags.GetTagsValue(Tag2);

                Evnt = "INT_EV_TOR_INSERT_WEIGHT_STR" + Convert.ToString(i).Trim();

                if (ActualCount[i] != PrevCount[i])
                {
                    PrevCount[i] = ActualCount[i];

                    Tags.UpdateTagsValue(Tag2, PrevCount[i]);
                    Tags.UpdateTagsValue(Evnt, 1);
                }
                else
                    Tags.UpdateTagsValue(Evnt, 0);

                // Checks the change of heat on Torchman
                Tag1 = "TOR_STR" + Convert.ToString(i) + "_HEAT_1";
                Tag2 = "TOR_STR" + Convert.ToString(i) + "_HEAT_2";

                AcHeatYear[i] = (int)Tags.GetTagsValue(Tag1);
                AcHeatNumber[i] = (int)Tags.GetTagsValue(Tag2);

                Tag1 = "TOR_STR" + Convert.ToString(i).Trim() + "_OPEN";
                isCasting = (bool)Tags.GetTagsValue(Tag1);

                // If AcHeatYear(i) * 10000 + AcHeatNumber(i) <> TorPrHeat And AcHeatYear(i) * 10000 + AcHeatNumber(i) <> TorAcHeat And TorPrHeat <> 0 Then
                // If (AcHeatYear(i) * 10000 + AcHeatNumber(i) <> TorPrHeat) And AcHeatYear(i) * 10000 + AcHeatNumber(i) <> TorAcHeat And IsCasting Then
                if (AcHeatYear[i] * 10000 + AcHeatNumber[i] > TorAcHeat && isCasting)
                    HeatChange = true;
            }

            if (HeatChange)
            {
                // Disparar El EOH y SOH en Torchman
                TOR_EOH(Tags);
                TOR_SOH(Tags);
            }
        }
        public void TOR_SOH(TagHandler Tags)
        {
            // Buscar el nuevo n�mero
            // Const MaxStrands = 4

            int HeatYear;
            int HeatNumber;

            int AcHeatYear;
            int AcHeatNumber;

            int NewHeatYear = 0;
            int NewHeatNumber = 0;

            int PrHeatYear;
            int PrHeatNumber;
            DateTime StartTime;

            //int MajYear;
            int MajNumber;
            int AuxHeat;
            int isCasting;
            //int HeatIndex;
            string tag;

            int i;

            MajNumber = 0;


            for (i = 1; i <= MaxStrands; i++)
            {
                tag = "TOR_STR" + Convert.ToString(i) + "_HEAT_1";
                HeatYear = (int)Tags.GetTagsValue(tag);

                tag = "TOR_STR" + Convert.ToString(i) + "_HEAT_2";
                HeatNumber = (int)Tags.GetTagsValue(tag);

                tag = "TOR_STR" + Convert.ToString(i) + "_OPEN";
                isCasting = (int)Tags.GetTagsValue(tag);

                AuxHeat = (HeatYear * 10000 + HeatNumber) * isCasting;

                if (MajNumber < AuxHeat)
                {
                    MajNumber = AuxHeat;
                    NewHeatYear = HeatYear;
                    NewHeatNumber = HeatNumber;
                }
            }

            // Actualizar nuevo n�mero y fecha
            if (MajNumber > 0)
            {
                AcHeatYear = (int)Tags.GetTagsValue("INT_TOR_ACTUAL_HEAT_1");
                AcHeatNumber = (int)Tags.GetTagsValue("INT_TOR_ACTUAL_HEAT_2");

                PrHeatYear = AcHeatYear;
                PrHeatNumber = AcHeatNumber;

                AcHeatYear = NewHeatYear;
                AcHeatNumber = NewHeatNumber;
                StartTime = DateTime.Now;

                Tags.UpdateTagsValue("INT_TOR_PREV_HEAT_1", PrHeatYear);
                Tags.UpdateTagsValue("INT_TOR_PREV_HEAT_2", PrHeatNumber);

                Tags.UpdateTagsValue("INT_TOR_ACTUAL_HEAT_START_TIME", StartTime);

                Tags.UpdateTagsValue("INT_TOR_ACTUAL_HEAT_1", AcHeatYear);
                Tags.UpdateTagsValue("INT_TOR_ACTUAL_HEAT_2", AcHeatNumber);

                // Disparar la inserci�n de SOH
                if (AcHeatNumber != 0)
                    Tags.UpdateTagsValue("INT_EV_TOR_SOH", 1);
            }
        }
        public void TOR_EOH(TagHandler Tags)
        {
            // Calcular la duraci�n
            DateTime StartTime;
            int HeatDuration;
            int MinutesDuration;
            int SecondsDuration;

            StartTime = (DateTime)Tags.GetTagsValue("INT_TOR_ACTUAL_HEAT_START_TIME");
            if (StartTime == null)
                StartTime = DateTime.Now;

            HeatDuration = (int)(StartTime - DateTime.Now).TotalSeconds;
            MinutesDuration = HeatDuration / 60;
            SecondsDuration = HeatDuration % 60;

            Tags.UpdateTagsValue("INT_TOR_ACTUAL_HEAT_DUR_MIN", MinutesDuration);
            Tags.UpdateTagsValue("INT_TOR_ACTUAL_HEAT_DUR_SEC", SecondsDuration);

            // Disparar la actualizaci�n
            Tags.UpdateTagsValue("INT_EV_TOR_EOH", 1);
        }
        private void AccumulateStrandTimePwrOn_Off(TagHandler Tags)
        {
            // Private Sub AccumulateStrandTimePwrOn_Off(ChangeStatus As Boolean)

            // ChangeStatus means: if is calling from the event of change the Power On/Off of the furnace
            // Or just from other procedures to accumulate times.
            // If is from the ChangeStatus, must accumulate to the previous status.

            // Const MaxStrands = 4

            DateTime LastTimeUpdate;
            //int PowerOn;
            int SecDiff;
            int CastingSec;
            int NotCastingSec;
            int isCasting;
            int i;
            string tag;
            int HeatNo;
            int AllStCasting;

            // First Verifies if there is a Heat on Caster, and if all the strand are casting
            HeatNo = (int)Tags.GetTagsValue("CCM_HEAT_2");
            AllStCasting = 0;
            for (i = 1; i <= MaxStrands; i++)
            {
                tag = "TOR_STR" + Convert.ToString(i).Trim() + "_OPEN";
                AllStCasting += (int)Tags.GetTagsValue(tag);
            }


            LastTimeUpdate = (DateTime)Tags.GetTagsValue("INT_CAST_LAST_CASTING_TIME");
            if (LastTimeUpdate == null)
                LastTimeUpdate = DateTime.Now;

            // Only increments the counters if The heat on caster is <>0 and at least one strand is casting.

            if ((AllStCasting > 0 | HeatNo > 0))
            {
                SecDiff = (int)(LastTimeUpdate - DateTime.Now).TotalSeconds;

                for (i = 1; i <= MaxStrands; i++)
                {
                    tag = "TOR_STR" + Convert.ToString(i).Trim() + "_OPEN";
                    isCasting = (int)Tags.GetTagsValue(tag);

                    tag = "INT_CAS_STR" + Convert.ToString(i).Trim() + "_CASTING_SEC";
                    CastingSec = (int)Tags.GetTagsValue(tag);

                    tag = "INT_CAS_STR" + Convert.ToString(i).Trim() + "_NOT_CASTING_SEC";
                    NotCastingSec = (int)Tags.GetTagsValue(tag);

                    // If (isCasting = 1 And Not (ChangeStatus)) Or (isCasting = 0 And ChangeStatus) Then
                    if (isCasting == 1)
                    {
                        CastingSec += SecDiff;
                        tag = "INT_CAS_STR" + Convert.ToString(i).Trim() + "_CASTING_SEC";
                        Tags.UpdateTagsValue(tag, CastingSec);
                    }
                    else
                    {
                        NotCastingSec += SecDiff;
                        tag = "INT_CAS_STR" + Convert.ToString(i).Trim() + "_NOT_CASTING_SEC";
                        Tags.UpdateTagsValue(tag, NotCastingSec);
                    }
                }
            }
            LastTimeUpdate = DateTime.Now;

            Tags.UpdateTagsValue("INT_CAST_LAST_CASTING_TIME", LastTimeUpdate);

        }
        private int StageIndex(object Stage)
        {
            int idx = 0;
            switch (Stage)
            {
                case 0: idx = 0; break;
                case 10: idx = 1; break;
                case 12: idx = 2; break;
                case 20: idx = 3; break;
                case 22: idx = 4; break;
                case 30: idx = 5; break;
                case 40: idx = 6; break;
                case 50: idx = 7; break;
                case 80: idx = 8; break;
            }

            return idx;
        }
        private void CalculateTimeTundish(TagHandler Tags)
        {
            int year;
            int month;
            int day;
            int hour;
            int minute;
            int second;
            double CalcDate;

            year = (int)Tags.GetTagsValue("CAS_TUNDISH_YEAR");
            month = (int)Tags.GetTagsValue("CAS_TUNDISH_MONTH");
            day = (int)Tags.GetTagsValue("CAS_TUNDISH_DAY");
            hour = (int)Tags.GetTagsValue("CAS_TUNDISH_HOUR");
            minute = (int)Tags.GetTagsValue("CAS_TUNDISH_MINUTE");
            second = (int)Tags.GetTagsValue("CAS_TUNDISH_SECOND");

            double dt1 = new DateTime(year, month, day).ToOADate();
            double dt2 = new DateTime(hour, minute, second).ToOADate();
            CalcDate = dt1 + dt2;

            Tags.UpdateTagsValue("INT_CAS_TUNDISH_LAST_CHANGE_TIME", CalcDate);
        }
        public void CallMillSPPlan()
        {
            // ---- manda derecho el stored a la pila
            string cmdSend;
            cmdSend = "exec ssp_get_plc_mill_plan_in 'B','0'";

            MQHandler queue = new MQHandler();
            queue.CreateQueue();
            queue.SendMsg(cmdSend);
        }
        public void WhiBMGapChange(TagHandler Tags)
        {
            // GAP_BM_DURATION_HR + GAP_BM_DURATION_MIN + GAP_BM_DURATION_SEC + GAP_BM_START_HR + GAP_BM_START_MIN + GAP_BM_START_SEC
            int DurationHr;
            int DurationMin;
            int DurationSec;
            int StartHr;
            int StartMin;
            int StartSec;
            int StartTrig;

            string StartTime;


            DurationHr = (int)Tags.GetTagsValue("GAP_BM_DURATION_HR");
            DurationMin = (int)Tags.GetTagsValue("GAP_BM_DURATION_MIN");
            DurationSec = (int)Tags.GetTagsValue("GAP_BM_DURATION_SEC");
            StartHr = (int)Tags.GetTagsValue("GAP_BM_START_HR");
            StartMin = (int)Tags.GetTagsValue("GAP_BM_START_MIN");
            StartSec = (int)Tags.GetTagsValue("GAP_BM_START_SEC");
            StartTrig = (int)Tags.GetTagsValue("GAP_BM_START_TRIG");


            StartTime = ("00" + Convert.ToString(StartHr).Trim()).Right(2) + ":" + ("00" + Convert.ToString(StartMin).Trim()).Right(2) + ":" + ("00" + Convert.ToString(StartSec).Trim()).Right(2);

            // ---- manda derecho el stored a la pila
            string cmdSend;
            cmdSend = "exec ssp_insert_plc_mill_gap 'B'," + Convert.ToString(StartTrig) + "," + "'" + StartTime + "'," + Convert.ToString(DurationHr).Trim() + "," + Convert.ToString(DurationMin).Trim() + "," + Convert.ToString(DurationSec).Trim();

            MQHandler queue = new MQHandler();
            queue.CreateQueue();
            queue.SendMsg(cmdSend);
        }
    }
}