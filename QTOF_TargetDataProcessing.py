import os
import pandas as pd
import numpy as np
import re

######## DEV NOTES #########
#make sure openpyxl is installed for users with pandas

### This is a translation of Conrad Pritchard's Target Data Processing Matlab script into Python
### This is an attempt at a direct-ish translation
### It should make an identical output to the Matlab script
### This may be reworked at a later date



### SECTION 1 ###

#GENERAL NOTES FOR USERS GO HERE

#Acronyms and Abbreviations
#   DB: Double Blank
#   MB: Method Blank
#   LB: Laboratory Blank
#   LCS: Laboratory Control Sample
#   LLLCS: Low Level Laboratory Control Sample
#   LCSD: Outdated, has been replaced by LCS/LLLCS; the script will still handle it though.
#   CCV: Continuing Calibration Verification
#   ISC: Instrument Sensitivity Check
#   IS: Internal Standard
#   InjS: Injection Standard
#   aQC: Analytical Quality Control
#   mQC: Method Quality Control
#   cal: calibration standards, i.e. your calibration curve
#   aQC: CCV, ISC, LB (analytical QC samples)
#       Continuing Calibration Verifications
#       Instrument Sensitivity Check
#       Laboratory Blank
#   mQC: LCS, LLLCS, LCSD, MB (method QC samples)
#       Laboratory Control Sample
#       Low-Level Laboratory Control Sample
#       LCSD?
#       Method Blank

#### NAMING RULES ####
#utilize the naming rules given in the excel sheets
#LIST THE NAMING RULES HERE
#for diluted samples, naming format should be _d##x, where ## is the dilution factor
#example, for a sample that is diluted 200x name it "tote1_d200x"

### SECTION 2 ###

##Inputs and Options:

#here you will type the file name (INCLUDING the filetype extension) exactly
file_name = 'aquagga_NEGtarget_test.txt'

#here you will type the path name. If you don't know how to do this: In Windows you can "Copy as Path." On Mac you right click with Option held or you right click then hold option before clicking "copy."
##Future goal: make it so you put the script in a top level folder and it cycles through all folders instead of requiring a hard input of the path name
path_name = r"C:\Users\richa\OneDrive\Documents\Research_Mines\Higgins_Group\Data_Processing_Script_Rework\testing"

# Subtract MB (Method Blank) values from samples to correct for background contamination
subMB = False # False = Off, True = On

# In case of small volume injection direct injection run
direct_inject = True #True for a direct injection, false if SPE method. By default, the previous script assumed all SVI runs to be SPE and all LVI runs to be Direct Inject.

# If you are doing an SPE method you need the samplemass.xlsx file as well
## IMPORTANT ## 
# Keep this file in the same folder as your data files, it will use the same path name!
# Make sure to have the sheet for your run named the same as your LCMS data filename without the file extension
# example: if my LCMS data file is named "20260424_ES_USCneg.txt"
# then my samplemass sheet name needs to be "20260424_ES_USCneg"
sample_mass_sheet = 'samplemass.xlsx'


### SECTION 3 ###

#This section will read the file into a Pandas data frame

#Here the script determines if the input is a text file or excel file (SCIEX outputs text, the orbitrap outputs Excel)
name, ext = os.path.splitext(file_name)
if ext not in (".txt", ".xlsx"):
    raise ValueError(f"Error: unsupported filetype {ext}, must be .txt or .xslx")


#this line will import the output file into pandas as a data frame
if ext == ".txt":
    df = pd.read_csv(os.path.join(path_name, file_name), sep="\t", dtype=str, keep_default_na=False)
elif ext == ".xlsx":
    df = pd.read_excel(os.path.join(path_name, file_name), dtype=str, keep_default_na=False)

## Checking column names within the data frame
# this is analogous to section 3 of the MatLab script, but pandas doesn't need to do section 3.
# this is just a check for readability

required_cols = ["Sample Name", "Sample Index", "Injection Volume", "Component Name", "Component Group Name", "Component Type", "Retention Time", "Precursor Mass", "Mass Error Confidence", "Mass Error (ppm)", "IS Name", "Area", "IS Area", "Actual Concentration", "Calculated Concentration", "Sample Type", "Used", "Polarity"]
for col in required_cols:
    if col not in df.columns:
        raise ValueError(f"Column '{col}' not found in data file")



### SECTION 4 ###

## Determine Analysis Method and QC's (Direct injection vs SPE)
#IMPORTANT, previous code differentiates method by injection volume. Make sure to specify whether or not your method is direct inject or SPE above if you ran a SVI Direct Inject!

method_map = {
    1000: "1mL",
    100: "100uL",
}

inj_vol = int(df["Injection Volume"].iloc[0])


if inj_vol not in method_map:
    raise ValueError (f"Error: unknown analysis method (injection volume {inj_vol} uL not recognized)")

polarity = df["Polarity"].iloc[1]

meth = method_map[inj_vol]

if meth == "100uL" and direct_inject:
    meth = "100uL_DI"

elif meth == "100uL" and polarity == "Negative":
    CCVactconc = 250 #CCV/QC concentration [ng/L]
    ISCactconc = 25  #ISC concentration [ng/L]
    EPAactconc = 250 #EPA Bullseye [ng/L] SHOULD be 250

elif meth == "100uL" and polarity == "Positive":
    CCVactconc = 200 #CCV/QC concentration [ng/L]
    ISCactconc = 10  #ISC concentration [ng/L]

elif meth == "1mL":
    CCVactconc = 200    #CCV/QC concentration [ng/L]
    ISCactconc = 33.33  #ISC concentration [ng/L]
    EPAactconc = 333.33 #EPA Bullseye [ng/L] 

elif meth == "100uL_DI":
    CCVactconc = 200    #CCV/QC concentration [ng/L]
    ISCactconc = 33.33  #ISC concentration [ng/L]
    EPAactconc = 333.33 #EPA Bullseye [ng/L] 




### SECTION 5 ###

#section 5 from matlab script not needed in python. That section was for handling an issue with Windows MatLab that is bypassed using Pandas.
#this section is analagous to section 6 from the matlab script

## Identify Number of Compounds and Samples
#builds a new data frame from just the first sample to extract compound names, precursor masses, component group name, and the Internal Standard assigned to that compound
first_sample = df[df["Sample Name"] == df["Sample Name"].iloc[0]]

targets = first_sample[
    (first_sample["Component Type"].isin(["Quantifiers", "Qualifiers"])) &
    (first_sample["Component Group Name"] != "Inj Stds")
]

comnamelist = targets["Component Name"].tolist()
commass     = targets["Precursor Mass"].tolist()
comgroup    = targets["Component Group Name"].tolist()
comISname   = targets["IS Name"].tolist()

if meth == "100uL":
    inj_stds = first_sample[first_sample["Component Group Name"] == "Inj Stds"]
    InjSnamelist = inj_stds["Component Name"].tolist()
else:
    InjSnamelist = []

numcom = len(comnamelist)
numIS = len(first_sample[
    first_sample["Component Type"].isin(["Internal Standard", "Internal Stardards"]) &
    (first_sample["Component Group Name"] != "Inj Stds")
])



### SECTION 6 ###
#analagous to section 7 of the matlab

#Data extraction step
#Categories:
#   DB: double blanks
#   cal: calibration standards
#   aQC: CCV, ISC, LB (analytical QC samples)
#       Continuing Calibration Verifications
#       Instrument Sensitivity Check
#       Laboratory Blank
#   mQC: LCS, LLLCS, LCSD, MB (method QC samples)
#       Laboratory Control Sample
#       Low-Level Laboratory Control Sample
#       LCSD?
#       Method Blank
#   Extract actual calibration concentration, actual IS conc, IS peak area,
#   measured concentration. If value is not a number, it is recorded as NaN.
#   Common non-numerical values include 'N/A', '<0', etc

##Extract Value function##

#def extractvalue(val):
    # if val is None:
    #     return np.nan
    # if isinstance(val, float) and np.isnan(val):
    #     return np.nan
    # if isinstance(val, str):
    #     stripped = val.strip()
    #     if stripped in ("N/A", "<0", ""):
    #         return np.nan
    #     try:
    #         return float(stripped)
    #     except ValueError:
    #         return np.nan
    # return float(val)

##Sorting Samples into data frames by Sample Type
is_cal = df["Sample Type"] == "Standard"
is_aQC = ~is_cal & df["Sample Name"].str.contains("CCV|ISC|LB|EPA|QC")
is_mQC = ~is_cal & ~is_aQC & df["Sample Name"].str.contains("LCS|LLLCS|LCSD|MB")
is_skip = ~is_cal & ~is_aQC & ~is_mQC & df["Sample Name"].str.contains("DB|blank check|AFFF")
is_sam = ~is_cal & ~is_aQC & ~is_mQC & ~is_skip

cal = df[is_cal]
aQC = df[is_aQC]
mQC = df[is_mQC]
sam = df[is_sam]

#This will allow us to label things as target, injection standard, or internal standard, so that we can separate between them.
is_target = (
    df["Component Type"].isin(["Quantifiers", "Qualifiers"]) &
    (df["Component Group Name"] != "Inj Stds")
)

is_inj_std = (
    (df["Component Group Name"] == "Inj Stds") &
    (meth == "100uL")
)

is_int_std = (
    df["Component Type"].isin(["Internal Standard", "Internal Standards"]) &
    (df["Component Group Name"] != "Inj Stds")
)

#Data frames that separate between targets and injection standards

cal_targets = df[is_cal & is_target]
aQC_targets = df[is_aQC & is_target]
mQC_targets = df[is_mQC & is_target]
sam_targets = df[is_sam & is_target]

# injection standards only needed for 100uL method
cal_inj  = df[is_cal & is_inj_std]
aQC_inj  = df[is_aQC & is_inj_std]
mQC_inj  = df[is_mQC & is_inj_std]
sam_inj  = df[is_sam & is_inj_std]

#Creating a matrix of compounds x samples
#Making a make_pivot function to facilitate this
#Reindex is to make sure that the rows are always in compound order

def make_pivot(df_subset, value_col, index_list=None):
    if index_list is None:
        index_list = comnamelist
    return df_subset.pivot_table(
        index="Component Name",
        columns="Sample Name",
        values=value_col,
        aggfunc="first"
    ).reindex(index_list)

#pd.tonumeric converts to numbers, because the cells are read in as strings (text) originally.

sam_PA      = make_pivot(sam_targets, "Area").apply(pd.to_numeric, errors="coerce") #Sample Peak Area
sam_conc    = make_pivot(sam_targets, "Calculated Concentration").apply(pd.to_numeric, errors="coerce") #Sample Calculated concentration
sam_ISPA    = make_pivot(sam_targets, "IS Area").apply(pd.to_numeric, errors="coerce") #Sample IS Peak Area
sam_ISact   = make_pivot(sam_targets, "IS Actual Concentration").apply(pd.to_numeric, errors="coerce") #Sample IS Actual Concentration
sam_rnum    = make_pivot(sam_targets, "Sample Index").apply(pd.to_numeric, errors="coerce") #Sample index
sam_inj_PA  = make_pivot(sam_inj, "Area", index_list=InjSnamelist).apply(pd.to_numeric, errors="coerce") #Injection Standard Peak Area

aQC_PA      = make_pivot(aQC_targets, "Area").apply(pd.to_numeric, errors="coerce") 
aQC_conc    = make_pivot(aQC_targets, "Calculated Concentration").apply(pd.to_numeric, errors="coerce")
aQC_ISPA    = make_pivot(aQC_targets, "IS Area").apply(pd.to_numeric, errors="coerce") 
aQC_ISact   = make_pivot(aQC_targets, "IS Actual Concentration").apply(pd.to_numeric, errors="coerce") 
aQC_rnum    = make_pivot(aQC_targets, "Sample Index").apply(pd.to_numeric, errors="coerce") 
aQC_inj_PA  = make_pivot(aQC_inj, "Area", index_list=InjSnamelist).apply(pd.to_numeric, errors="coerce") 

mQC_PA      = make_pivot(mQC_targets, "Area").apply(pd.to_numeric, errors="coerce") 
mQC_conc    = make_pivot(mQC_targets, "Calculated Concentration").apply(pd.to_numeric, errors="coerce")
mQC_ISPA    = make_pivot(mQC_targets, "IS Area").apply(pd.to_numeric, errors="coerce") 
mQC_ISact   = make_pivot(mQC_targets, "IS Actual Concentration").apply(pd.to_numeric, errors="coerce") 
mQC_rnum    = make_pivot(mQC_targets, "Sample Index").apply(pd.to_numeric, errors="coerce") 
mQC_inj_PA  = make_pivot(mQC_inj, "Area", index_list=InjSnamelist).apply(pd.to_numeric, errors="coerce")

cal_PA      = make_pivot(cal_targets, "Area").apply(pd.to_numeric, errors="coerce") 
cal_conc    = make_pivot(cal_targets, "Calculated Concentration").apply(pd.to_numeric, errors="coerce")
cal_actconc = make_pivot(cal_targets, "Actual Concentration").apply(pd.to_numeric, errors="coerce")
cal_ISPA    = make_pivot(cal_targets, "IS Area").apply(pd.to_numeric, errors="coerce") 
cal_ISact   = make_pivot(cal_targets, "IS Actual Concentration").apply(pd.to_numeric, errors="coerce") 
cal_rnum    = make_pivot(cal_targets, "Sample Index").apply(pd.to_numeric, errors="coerce") 
cal_used    = make_pivot(cal_targets, "Used")
cal_inj_PA  = make_pivot(cal_inj, "Area", index_list=InjSnamelist).apply(pd.to_numeric, errors="coerce")


#MB subtraction from matlab script isn't actually handled in this section despite there being something written. I am not including that stub here.
#MB subtraction is handled in a later section.


### SECTION 7 ###
#analagous to section 8 of the matlab

##Determining the calibration range

#Sorting cal columns by actual concentration. This just uses the first compound since the cal curve sort is the same for all compounds.

sort_order   = cal_actconc.loc[comnamelist[0]].apply(pd.to_numeric, errors="coerce").argsort()
cal_actconc_s = cal_actconc.iloc[:, sort_order]
cal_conc_s    = cal_conc.iloc[:, sort_order]
cal_used_s    = cal_used.iloc[:, sort_order]

lowlim = {}
highlim = {}

for compound in comnamelist:
    valid = []
    for col in cal_actconc_s.columns:
        actconc = pd.to_numeric(cal_actconc_s.loc[compound, col], errors="coerce")
        conc    = pd.to_numeric(cal_conc_s.loc[compound, col], errors="coerce")
        used    = str(cal_used_s.loc[compound, col]).strip().upper()
        acc     = conc / actconc if actconc !=0 else np.nan
        if 0.7 < acc < 1.3 and used == "TRUE":
            valid.append(actconc)
    
    lowlim[compound]    = valid[0] if valid else np.nan
    highlim[compound]   = valid[-1] if valid else np.nan

### SECTION 8 ###
#analagous to section 9 of the matlab

## Determining IS Peak Area Recovery

#SVI Negative method
if meth == "100uL" and polarity == "Negative":
    #This will read the samplemass excel file and makes a dataframe with the two columns from the NIS sheet
    NIS = pd.read_excel(
        os.path.join(path_name, "samplemass.xlsx"),
        sheet_name="NIS",
        header=0,
        dtype=str
    )
    #this makes a dictionary that maps the compoudn name to the injection std name
    #e.g. {"PFOA":"M8PFOA", "PFOS":"M8PFOS", etc}
    NIS_dict = dict(zip(NIS.iloc[:, 0], NIS.iloc[:, 1]))
    # builds list of which injection standard each compound uses
    # also builds a list of compoudns missing from the NIS sheet
    InjSlist = []
    missing = []
    for compound in comnamelist:
        if compound not in NIS_dict:
            missing.append(compound)
        else:
            InjSlist.append(NIS_dict[compound])
    if missing:
        raise ValueError(f"The following compounds were not found in the NIS sheet of samplemass.xlsx:\n" + "\n".join(missing) )

    # builds matched injection standard PA matrices aligned to comnamelist
    cal_injmatch = cal_inj_PA.loc[[NIS_dict.get(c, InjSnamelist[0]) for c in comnamelist]]
    cal_injmatch.index = comnamelist

    aQC_injmatch = aQC_inj_PA.loc[[NIS_dict.get(c, InjSnamelist[0]) for c in comnamelist]]
    aQC_injmatch.index = comnamelist

    mQC_injmatch = mQC_inj_PA.loc[[NIS_dict.get(c, InjSnamelist[0]) for c in comnamelist]]
    mQC_injmatch.index = comnamelist

    sam_injmatch = sam_inj_PA.loc[[NIS_dict.get(c, InjSnamelist[0]) for c in comnamelist]]
    sam_injmatch.index = comnamelist

    # calculates the mean cal IS PA and IS actual concentration for recovery normalization
    cal_ISPA_mean  = cal_ISPA.mean(axis=1)
    cal_ISact_mean = cal_ISact.mean(axis=1)
    cal_inj_mean   = cal_injmatch.mean(axis=1)

    # IS recovery
    sam_ISrec = (sam_ISPA / sam_injmatch) / (cal_ISPA_mean / cal_inj_mean).values[:, None] * (cal_ISact_mean / sam_ISact).values[:, None]
    aQC_ISrec = (aQC_ISPA / aQC_injmatch) / (cal_ISPA_mean / cal_inj_mean).values[:, None] * (cal_ISact_mean / aQC_ISact).values[:, None]
    mQC_ISrec = (mQC_ISPA / mQC_injmatch) / (cal_ISPA_mean / cal_inj_mean).values[:, None] * (cal_ISact_mean / mQC_ISact).values[:, None]

    # injection standard recovery
    sam_injrec = sam_inj_PA / cal_inj_PA.mean(axis=1).values[:, None]
    aQC_injrec = aQC_inj_PA / cal_inj_PA.mean(axis=1).values[:, None]
    mQC_injrec = mQC_inj_PA / cal_inj_PA.mean(axis=1).values[:, None]

else:
    InjSlist = []
    cal_ISPA_mean  = cal_ISPA.mean(axis=1)
    cal_ISact_mean = cal_ISact.mean(axis=1)

    sam_ISrec = sam_ISPA / cal_ISPA_mean.values[:, None] * (cal_ISact_mean.values[:, None] / sam_ISact)
    aQC_ISrec = aQC_ISPA / cal_ISPA_mean.values[:, None] * (cal_ISact_mean.values[:, None] / aQC_ISact)
    mQC_ISrec = mQC_ISPA / cal_ISPA_mean.values[:, None] * (cal_ISact_mean.values[:, None] / mQC_ISact)

    sam_injrec = aQC_injrec = mQC_injrec = None

### SECTION 9 ###
#Check Analytical Accuracy

# identify which aQC columns belong to which QC type
aQC_cols = aQC_conc.columns

CCV_cols = [c for c in aQC_cols if "CCV" in c]
QC_cols  = [c for c in aQC_cols if "QC"  in c]
ISC_cols = [c for c in aQC_cols if "ISC" in c]
LB_cols  = [c for c in aQC_cols if "LB"  in c]
EPA_cols = [c for c in aQC_cols if "EPA" in c]

# calculate recoveries
# CCVactconc is used for both CCV and QC becaues 
# they're the same standard at the same concentration

CCVrec = aQC_conc[CCV_cols] / CCVactconc if CCV_cols else pd.DataFrame()
QCrec  = aQC_conc[QC_cols]  / CCVactconc if QC_cols  else pd.DataFrame()
ISCrec = aQC_conc[ISC_cols] / ISCactconc if ISC_cols else pd.DataFrame()
LBconc = aQC_conc[LB_cols]               if LB_cols  else pd.DataFrame()
EPArec = aQC_conc[EPA_cols] / EPAactconc if EPA_cols and EPAactconc else pd.DataFrame()

# combine into one matrix
aQCall     = pd.concat([CCVrec, QCrec, ISCrec, LBconc, EPArec], axis=1)
aQCallname = CCV_cols + QC_cols + ISC_cols + LB_cols + EPA_cols


### SECTION 10 ###
#Check method accuracy and determine reporting limits

RL = pd.Series(lowlim, index=comnamelist)  #Reporting limit. The RL starts out with the lower calibration limit determined earlier, and increases if the MB concentration exceeds it.

if len(mQC_targets) > 0:
    #if there are method QCs
    mQC_cols = mQC_conc.columns
    #checks sample names for LCS, LLLCS, LCSD, and MB
    LCS_cols = [c for c in mQC_cols if "LCS" in c]
    MB_cols  = [c for c in mQC_cols if "MB" in c]

    LCSconc = mQC_conc[LCS_cols] if LCS_cols else pd.DataFrame()
    MBconc  = mQC_conc[MB_cols]  if MB_cols  else pd.DataFrame()

    if MB_cols:
        MB_mean = MBconc.mean(axis=1) #calculates the MB mean concentration across columns

        if subMB: #if you are subtracting MB concentrations from the background to correct for background contamination
            #this option is specified in the inputs section
            sam_conc_notMBcorr = sam_conc.copy() #this saves a copy of the uncorrected concentrations before the MB subtraction
            #the uncorrected concentrations will be in the output file
            sam_conc = sam_conc.subtract(MB_mean, axis=0)
        # sets the reporting limit (RL) to the max MB concentration or to lowlim, whichever is higher
        for compound in comnamelist:
            MB_max = MBconc.loc[compound].max()
            if subMB:
                MB_max -= MB_mean[compound]
            RL[compound] = MB_max if MB_max > lowlim[compound] else lowlim [compound]

    #combines things for the output file
    subMBval    = MB_mean if subMB and MB_cols else pd.Series(0, index=comnamelist) #establishes the subMB value as the Mean MB if subMB is True. 
    #subMBval is used for calculating Matrix concs
    mQCall      = pd.concat([LCSconc, MBconc], axis=1) #this combines the LCS and MB dataframes together side by side
    mQCallname  = LCS_cols + MB_cols

else: #if there are no method QCs. The reporting limit will stay the lower calibration limit (lowlim)
    mQCall              = pd.DataFrame()
    mQCallname          = []
    sam_conc_notMBcorr  = None


### SECTION 11 ###
##Calculates Matrix Concentration, Matrix LOQ range, and Replaces values outside Matrix LOQ

if meth == "100uL":   #if doing SVI with SPE
    soildata = pd.read_excel(
        os.path.join(path_name, "samplemass.xlsx"), #reads the samplemass excel, make sure it is in the same directory/folder as your data!
        sheet_name=name,
        header=0,
        dtype=str
    )

    #Extract the relevant columns from the samplemass sheet
    #Converts wetmass and moisture frac to numbers since the excel was read in as strings
    sample_names    = soildata.iloc[:, 0]
    wet_mass        = pd.to_numeric(soildata.iloc[:, 1], errors="coerce")
    moisture_frac   = pd.to_numeric(soildata.iloc[:, 2], errors="coerce")
    matrix_type     = soildata.iloc[:, 3]

    dry_mass    = wet_mass * (1 - moisture_frac) #determines dry mass from moisture fraction and wet mass
    sd_idx      = match.index[0]
    mtype       = matrix_type[sd_idx].strip()
    mass        = dry_mass[sd_idx]

    if pd.isna(mass) or mass == 0:
        raise ValueError(f"Sample '{sample}' has a miss or zero dry mass in samplemass.xlsx")

    # extraction factors per matrix type
    ext_factors = {
        "soil"      : (0.4 / 0.1)  * (1.5 / 1000),
        "dust"      : (0.6 / 0.45) * (1.5 / 1000) * (5 / 0.25),
        "water"     : (0.75 / 0.375) * (5 / 1000) * (1000),
        "SPEsoil"   : (0.75 / 0.375) * (5 / 1000) * (32 / 2.5),
        "Tissue"   : (0.75 / 0.375) * (5 / 1000) * (10 / 2.5)
                    # (total vol / sample vol in vial) * (total extract vol mL / 1000 converts to L) * (pre-SPE extraction vol / volume of pre-SPE put through SPE)
    }

    # builds a matrix concentration dataframe
    mat_samconc = pd.DataFrame(index=comnamelist, columns=sam_conc.columns, dtype=float)
    mat_highlim = pd.DataFrame(index=comnamelist, columns=sam_conc.columns, dtype=float)
    mat_lowlim  = pd.DataFrame(index=comnamelist, columns=sam_conc.columns, dtype=float)
    
    for i, sample in enumerate(sam_conc.columns):
        #matches the sample name to soildata
        match = soildata[sample_names == sample]
        if match.empty:
            raise ValueError(f"Sample '{sample}' not found in samplemass.xlsx")
        
        sd_idx  = match.index[0]
        mtype   = matrix_type[sd_idx].strip()

        if mtype not in ext_factors:
            raise ValueError(f"Unknown matrix type '{mtype}' for sample '{sample}'")
        
        ext = ext_factors[mtype]
        mass = dry_mass[sd_idx]

        mat_samconc.iloc[:, i]  = sam_conc.iloc[:, i] * ext / mass # matrix sample conc is sample conc * extraction factors for that matrix type
        mat_highlim.iloc[:, i]  = pd.Series(highlim, index=comnamelist) * ext / mass #matrix high limit also adjusted for "" matrix type
        mat_lowlim.iloc[:, i]   = RL * ext / mass #low lim is Reporting limit and adjusting for "" matrix type

    #determines the units based on matrix type
    units = "ng/L" if matrix_type.iloc[-1].strip() == "water" else "ng/g" #water gets ng/L, solids get ng/g

else:
    factor = 1.5 / 0.9 #factor takes into account the sample being from the water of the sample prep
    # The water fraction is 60%, 0.9mL is 60% of 1.5mL
    # for whatever reason, this doesn't include the basic water fraction ???
    mat_samconc = sam_conc * factor
    mat_highlim = pd.DataFrame(
        {col: pd.Series(highlim, index = comnamelist) * factor for col in sam_conc.columns}
    )
    mat_lowlim = pd.DataFrame(
        {col: RL * factor for col in sam_conc.columns}
    )
    units = "ng/L"

# replace values outside LoQ with < or > notation
#mat_concLOQ stays numeric throughout
mat_concLoQ = mat_samconc.copy().astype(object)

# values outside LOQ get replaced with the "<" or ">" limit value, flagged separetly in a new "flag" data frame
mat_flag = pd.DataFrame("", index=comnamelist, columns=sam_conc.columns)
for compound in comnamelist:
    for sample in sam_conc.columns:
        val  = mat_samconc.loc[compound, sample]
        high = mat_highlim.loc[compound, sample]
        low  = mat_lowlim.loc[compound, sample]
        #storing raw floats instead of rounded strings here, will round later to better match previous script
        if pd.isna(val) or val < low:
            mat_concLoQ.loc[compound, sample] = low
            mat_flag.loc[compound, sample] = "<" #saving the matrix conc of LOQ as numeric and making a new dataframe for the formatted version with "<" 
        elif val > high:
            mat_concLoQ.loc[compound, sample] = high
            mat_flag.loc[compound, sample] = ">" 
        else:
            mat_concLoQ.loc[compound, sample] = val #val is within range, flag stays "", value stays as-is

### SECTION 12 ###
## Corrects for Dilutions ##
#Make sure you are using "_d##x" for your sample naming
#Your dilution needs to be an integer, it won't handle decimals. This shouldn't be an issue with the current sample prep dilutions.

mat_concLoQ_dil = mat_concLoQ.copy()
mat_flag_dil    = mat_flag.copy()

#makes a copy of mat_concLoQ and mat_flag_dil to dilution correct while keeping the original mat_concLoQ

for sample in sam_conc.columns:
    dil_match = re.search(r"_d(\d+)x", sample, re.IGNORECASE) #finds your dilution factor, and if there is one. #IGNORECASE is in case you used _D##x or _d##X or some permutation thereof
    if dil_match:
        ndil = float(dil_match.group(1))
        mat_concLoQ_dil[sample] = mat_concLoQ[sample] * ndil
        #flags carry over unchanged, don't need to adjust here

# final formatted string dataframe for output
mat_concLoQ_final = mat_concLoQ_dil.copy().astype(object)

for compound in comnamelist:
    for sample in mat_concLoQ_dil.columns:
        flag = mat_flag_dil.loc[compound, sample]
        val = mat_concLoQ_dil.loc[compound, sample]
        if flag == "<":
            mat_concLoQ_final.loc[compound, sample] = f"<{val:.2f}"
        elif flag == ">":
            mat_concLoQ_final.loc[compound, sample] = f">{val:.2f}"
        else:
            mat_concLoQ_final.loc[compound, sample] = val



### SECTION 13 ###
## Writes all the data into an excel file ##

out_path = os.path.join(path_name, f"{name}_results_pytest.xlsx") #writes excel file to the same folder as the lcms data file

with pd.ExcelWriter(out_path, engine="openpyxl") as writer:

    # --- Tab 1: Sample Concentrations -------------
    sam_conc_out = mat_concLoQ_final.copy()
    sam_conc_out.insert(0, "Group", pd.Series(dict(zip(comnamelist, comgroup))))
    sam_conc_out.insert(0, "Precursor Mass", pd.Series(dict(zip(comnamelist, commass))))
    sam_conc_out.to_excel(writer, sheet_name="Sample Conc", index=True)

    # --- Tab 2: Surrogate Recovery ----------------
    sam_ISrec.to_excel(writer, sheet_name="Surrogate Recovery", index=True)

    # --- Tab 3: Method Accuracy -------------------
    if len(mQC_targets) > 0:
        mQCall.to_excel(writer, sheet_name="Method Accuracy", index=True)
    else:
        pd.DataFrame(["***NO METHOD ACCURACY SAMPLES FOUND***"]).to_excel(
            writer, sheet_name="Method Accuracy", index=False, header=False
        )
    
    # --- Tab 4: Analytical Accuracy --------------
    if len(aQC_targets) > 0:
        aQCall.to_excel(writer, sheet_name="Analytical Accuracy", index=True)
    else:
        pd.DataFrame(["***NO ANALYTICAL ACCURACY SAMPLES FOUND***"]).to_excel(
            writer, sheet_name="Analytical Accuracy", index=False, header=False
        )
    
    # --- Tab 5: IS and Quant Limits -------------
    IS_lims = pd.DataFrame({
        "Compound"                                          : comnamelist,
        "Internal Standard"                                 : comISname,
        "Injection Standard"                                : InjSlist if InjSlist else [""] * numcom,
        "Analytical Low Quant Lim [ng/L]"                   : [lowlim[c] for c in comnamelist],
        "Analytical High Quant Lim [ng/L]"                  : [highlim[c] for c in comnamelist],
        f"Matrix Low Quant Lim [{units}]"                   : mat_lowlim.min(axis=1).values,
        f"Matrix High Quant Lim [{units}]"                  : mat_highlim.max(axis=1).values,
    })
    IS_lims.to_excel(writer, sheet_name="IS and Quant Lims", index=False)

    # --- Tab 6: QA QC Notes ----------------------
    notes = pd.DataFrame([
        [f"All sample concentrations are reported in [{units}]"],
        ["Min LOQ set to 3x method blank or minimum calibration curve point, whichever is higher"],
        ["Continuing calibration sample recoveries are within 70%-130%, except:"],
        ["Instrument sensitivity check recoveries are within 70%-130%, except:"],
        ["Laboratory control sample recoveries are within 70%-130%, except:"],
        ["Cells highlighted yellow if surrogate recovery was outside 50-150%"],
        ["Cells highlighted red if surrogate recovery was outside 30-200%"],
        ["Additional Notes:"],
    ])
    notes.to_excel(writer, sheet_name="QA QC Notes", index=False, header=False)

    # --- Tab 7: Raw Vial Concentrations ----------
    sam_conc.to_excel(writer, sheet_name="Raw Vial Conc", index=True)

    # --- Tab 8: Raw Matrix Concentrations --------
    mat_samconc.to_excel(writer, sheet_name="Raw Corrected Conc", index=True)

    # --- Tab 9: Peak Areas
    # combine the sample indices and sort them for tabs 9 and 10

    all_rnum = pd.concat([cal_rnum, aQC_rnum, mQC_rnum, sam_rnum], axis=1)
    sort_order = all_rnum.iloc[0].argsort()

    all_PA = pd.concat([cal_PA, aQC_PA, mQC_PA, sam_PA], axis=1).iloc[:, sort_order]
    all_PA.to_excel(writer, sheet_name="Peak Area", index=True)

    # --- Tab 10: IS Peak Areas -------------------
    all_ISPA = pd.concat([cal_ISPA, aQC_ISPA, mQC_ISPA, sam_ISPA], axis=1).iloc[:, sort_order]
    all_ISPA.to_excel(writer, sheet_name="IS Peak Area", index=True)

    # --- Tab 11: Non-MB-corrected vial concentrations (optional) ------
    if subMB and sam_conc_notMBcorr is not None:
        sam_conc_notMBcorr.to_excel(
            writer, sheet_name="Original Raw Vial Conc", index=True
        )


print(f"Done. Output written to: {out_path}")