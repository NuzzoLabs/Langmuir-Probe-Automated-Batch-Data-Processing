# -*- coding: utf-8 -*-
"""
Created on Sat Jan  18 2025 

Langmuir Probe Batch Processing
Version 2.4

Same as version 2.3, except with bug fixes incorporated from testing of the Dave Pace data set.

NOTE: This version is only created for any potential modifications to existing functions due to the need to iterate steps 4-10 in Lobbia. 
Original functions that have been tested and verified to work are in version 2.0. This version only serves as a prototype for the final
code with the iterative steps implemented

Description: This program is used to process raw Langmuir Probe (LP) data files with current and voltage values. It takes a folder with all
the LP data files desired to be processed as input and automatically calculates the plasma paramters of all LP data in the folder. Once it
processes all the files it will generate an excel file with all the post-processed plasma parameters for each file.

This is the second version of this code. This version is more generalized to handle a wider range of LP data files and has the additions of 
electron number density and EEDF calculations (STILL NEED TO ADD). It is also more inline with Lobbia's methods. Especially those used for calculating 
ion and electron saturation current for calculating plasma number densities. This version provides more utility over version 1 and improved 
plasma properties analytics. 

Major Changes:
    
Version 2.6

Significant updates made in this version due to inefficiencies when processing large quantities of files. Multiple weaknesses were identified. Covered below.

1. Consolidated the plasma potential function into one function. Instead of having a separate function for the second derivative method and the Lobbia first derivative
   method, this was consolidated into one function that determines when to use one or the other. It will first try the preferred second derivative method and if that fails 
   it will default to Lobbia's first derivative method.
   
2. Removal of files with complex solutions to floating potential. In processing some of the large folders in the 2025 data set of LP data some files had asymptotic behavior near the current 
   axis (y-axis). The script calculates floating potential by fitting a quadratic equation to a few data points around the voltage axis (x-axis). This could be resolved by simply doing a linear
   regression, however, using the quadratic fit actually helped identify poor data sets since this asymptotic behavior lead to complex solutions for floating potential, meaning the quadratic fit
   never intersected the x-axis which is how floating potential is defined. This update identifies these complex solutions and then removes the files from the folder and resizes the floating potential
   array as well as the voltage and current arrays so that the bad files don't get processed. NOTE: IF YOU WANT TO KEEP ALL FILES FOR RECORD, IT IS RECOMMENDED TO BACKUP YOUR FILES IN A SEPARATE FOLDER 
   SINCE THE BAD ONES WILL BE DELETED BY THE SCRIPT! 
   
3. Floating potential function updated to order solutions based on concavity. In the original version of this function, floating potential was solved by taking the second solution to the quadratic equation
   with the assumption that the curve was concave up (i.e. positive slope going to the right of the graph). There are a few instances where the curve is concave down which changes the ordering of the solution
   since the solution is sorted from least to greatest. So the script checks if the curve is concave up or down and then sorts from least to greatest. Floating potential is whichever side of the parabola is
   uptrending since the LP curve is uptrending in the region where it crosses the x-axis.
   
4. Added a routine for removing files where Vp < Vf. Floating potential should never be greater than plasma potential so if this is the case there is something wrong with the data. This may cause the script to
   calculate a negative electron temperature if it defaults to the potential method which is not possible. Sometimes this may be adjustable by changing the window size in the plasma potential calculation, but
   it is easier to just remove these files when dealing with large amounts of data.
   
(TBD) 5. Consider adding a function that takes all files removed and places them in a specified location so that the analyst can look at the data to see if there are any obvious issues with the LP curve.
    
Version 2.5 

1. Updated ion saturation current calculation in the PlasmaProps() function to use a try/except architecture. This is to address the cases where
   the ion saturation region does not reach stabilization. If it fails the stabilization criteria in the original "try" block, it defaults to calculating
   the average of the ion current from the floating potential to the beginning of the file. If the except block is executed it indicates there is a problem
   with the ion saturation region as it isn't reaching stabilization which implies the probe voltage is probably needs to be swept to more negative values.


Version 2.4

1. Bug fixes from issues uncovered during code validation using the David Pace data set.
 1. a. In plasma properties function x and y values have a filter applied to only model data where current is >= 0 since for certain cases the 1/B exponent 
       needed to solve for plasma number density can yield undefined resulting in NaN values for negative currents.
 1. b. In the CL_Sheath function it was noticed that the probe area variable, Ap, was undefined in the function for some reason. Ap has been included as an input to 
       this function now. Also the option to have shape as an input was added to handle the case of planar probes where thin sheath is generally a good assumption.
       
Version 2.3
1. Bug fixes incorporated to get a functional iterative method
2. Solution convergence checks added and capability to write results to excel file

This program was validated on the data set "Sample Langmuir Data" that can be downloaded from the following:
https://davidpace.com/example-of-langmuir-probe-analysis/#ne

From Version 2.2 ***************Things that need looking into or improvement in next version****************
1. The method of determining ion saturation current needs to be checked.
    From the information that can be found it does not appear that there is a standard method so the averaging method in the
    stabilized region of where ion saturation should occur was kept for this calculation. Ideal LP traces will have the ion current
    completely saturate, but all data files used to test this code never displayed full saturation (demonstrating linear behavior in the region instead), 
    probably because the probe bias needed to be biased more negative.
    
2. The way that the code calculates ion number density might need to be revised. The A and B parameters are calculated in the iteration loop
and also calculated in the plasma props function where ion number density is calculated. The ion number density in the iterative method should
probably use the A and B parameters calculated in the loop, not inside the plasma properties function.
    It is still not clear why these ion number density is so small (like in the 10^-15 m^-3 range). These should be close to the electron number density.
    Currently electron number density is being used as the value for the plasma density since the amount of electrons are a pretty good estimation of how
    many charged particles are within the plasma. It would be interesting to look into this discrepancy more since the assumption that the plasma is quasi-neutral
    does not hold if there is this large of a discrepancy between the ion number density and the electron number density. Logically, the ion number density has to be 
    AT LEAST as big as the electron number density since for each electron ejected by ionization of each atom, there is one ion. If there is secondary ionization occuring
    then there should be more electrons than ions, but the number of ions has to be at least as big as the number of electrons. This math still needs to be checked.
    

    Version 2.7
    This version includes a new function to solve the Electron Energy Distribution Function (EEDF) of the plasma

Nomenclature
... This section to come at a later time once the program is finished.
Need to standardize all the variables used in the code to make it more readable to the user...

@author: Nick Nuzzo
"""

#!/usr/bin/env python
from decimal import Decimal
from statistics import mean
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import os
import glob
import math
try:
    import xlwt
    from xlwt import Workbook
except ImportError:  # Optional legacy Excel writer; CLI writes .xlsx with pandas.
    xlwt = None
    Workbook = None
from sklearn.metrics import r2_score
from scipy.signal import savgol_filter
from scipy.special import logsumexp
import statistics
import shutil
from scipy.integrate import odeint

# The following class is a class of the langmuir probe and its attributes that will be necessary for different parts of the analysis. This allows ease of
# implementation for certain steps that require certain LP attributes and repeat usage without needing specialized code for the specific condition or LP geometry
    
class LP:

    def __init__(self, shape):
        self.shape = shape
    
    # Method for calculating the parameters a and b that are necessary for correction to the ion density calculation per Lobbia equations 6, 7, and 8.
    def Params(self, rp_d):
		
        # For the transitional sheath case. See Lobbia equation 7 and 8
        if rp_d < 50 and rp_d > 3:     
            
            if self.shape == 'cylindrical':
                a = 1.18-0.0008*(rp_d)**1.35
                b = 0.0684 + (0.722+0.928*rp_d)**-0.729
                
            if self.shape == 'spherical':
                a = 1.58 + (-0.056+0.816*rp_d)**-0.744
                b = -0.933 + (0.0148+0.119*rp_d)**-0.125
                
            if self.shape == 'planar':
                a = np.exp(-1/2)*np.sqrt(2*math.pi)*(2.28*rp_d**-0.749)
                b = 0.806*rp_d**-0.0692
        
        # For the Orbital Motion Limited case (thick sheath). See Lobbia section F. Algorithm for Langmuir Probe Analysis (Single Probe) step 10c
        if rp_d <= 3:
            
            if self.shape == 'cylindrical':
                a = 2*np.sqrt(math.pi)
                b = 1/2
                
            if self.shape == 'spherical' or self.shape == 'planar':
                a = 1
                b = 1
        
        # For thin sheath
        if rp_d >= 50:
            return 'Correction not required. Thin-sheath assumption applies thus correction is unnecessary'
        
   
        return a, b
    
    # LP geometry-dependent ion current correction
    # As a general rule, use the electron density for "N0" in these equations. It is a more accurate representation of the overall plasma density 
    # unless there is significant secondary or tertiary ionizations.
    # This function will also return an array of voltage differences to be converted into the associated bias voltages for each corrected ion current values
    
    def IonCurrCor(self, rp_d, e, N0, Ap, mi, Te, differences, a, b):
        
        # For the transitional sheath case. See Lobbia equation 7
        if rp_d < 50 and rp_d > 3:
            if self.shape == 'cylindrical' or self.shape == 'spherical':
                

                # Calculate I using equation 
                I = [e * float(Ap) * float(N0) * np.sqrt((e * float(Te))/(2*math.pi*mi)) * a * (float(diff)/float(Te))**b for idx, diff in enumerate(differences) if differences > 0]
                
                # Creating mask for (Vp-V)/Te > 1 condition that must be true for the above equation to be applied. 
                
                mask = [np.divide(diff,Te) > 1 for idx, diff in enumerate(differences)]
                
                # Creating array of ion current where the above condition applies
                
                IC = np.array([I[idx][x] for idx, x in enumerate(mask) if x == True])
                
                # Creating array of voltage differences
                
                VCDiff = np.array([differences[idx][x] for idx, x in enumerate(mask) if x == True])
                
        # For the transitional sheath case with planar probe. See Lobbia equation 8    
        if rp_d < 45 and rp_d > 10:
            if self.shape == 'planar':
        
                # Calculate I using equation 
                I = [e * float(Ap) * float(N0) * np.sqrt((e * float(Te))/(2*math.pi*mi)) * a * (float(diff)/float(Te))**b + np.exp(-1/2)*e * Ap * float(N0)*np.sqrt(e*float(Te)/mi) for idx, diff in enumerate(differences) if differences > 0]
          
                # Creating mask for 3 < (Vp-V)/Te < 30 condition that must be true for the above equation to be applied. 
                
                mask = [(np.divide(diff,Te) > 3 and np.divide(diff,Te) < 30) for idx, diff in enumerate(differences)]
                
                # Creating array of ion current where the above condition applies
                
                IC = np.array([I[idx][x] for idx, x in enumerate(mask) if x == True])
                
                # Creating array of voltage differences
                
                VCDiff = np.array([differences[idx][x] for idx, x in enumerate(mask) if x == True])
                
        # For the Orbital Motion Limited case (thick sheath). See Lobbia equation 6
        if rp_d <= 3:
            if self.shape == 'cylindrical':

                # Calculate I using equation 
                I = [np.divide((e * float(Ap) * float(N0)), math.pi) * np.sqrt(np.divide((2 * e * float(diff)), mi)) for idx, diff in enumerate(differences) if diff > 0]
            
                # Creating mask for (Vp-V)/Te >> 1 condition that must be true for the above equation to be applied. Using 5 as the threshold, but this
                # can be changed depending on the results. ">> 1" is what is cited in Lobbia and is rather arbitrary.
                
                mask = [np.divide(diff,Te) > 5 for idx, diff in enumerate(differences) if differences[idx] > 0]
                
                # Creating array of ion current where the above condition applies
                
                IC = np.array([I[idx][x] for idx, x in enumerate(mask) if x == True])
                
                # Creating array of voltage differences
                
                VCDiff = np.array([differences[idx][x] for idx, x in enumerate(mask) if x == True])
            
            if self.shape == 'spherical' or self.shape == 'planar':

                # Calculate I using equation 
                I = [e * float(Ap) * float(N0) * np.sqrt((e * float(Te))/(2*math.pi*mi)) * float(diff)/float(Te) for idx, diff in enumerate(differences) if differences[idx] > 0]
                
                # Creating mask for (Vp-V)/Te >> 1 condition that must be true for the above equation to be applied. Using 5 as the threshold, but this
                # can be changed depending on the results. ">> 1" is what is cited in Lobbia and is rather arbitrary.
                
                mask = [np.divide(diff,Te)[idx] > 10 for idx, diff in enumerate(differences)]
                
                # Creating array of ion current where the above condition applies
                
                IC = np.array([I[idx][x] for idx, x in enumerate(mask) if x == True])
                
                # Creating array of voltage differences
                
                VCDiff = np.array([differences[idx][x] for idx, x in enumerate(mask) if x == True])
                
        return IC, VCDiff

####################################################### BEGINNING OF FUNCTIONS #################################################################

# The following function removes bad data files from the folder. Currently it only considers files that have no positive current values as
# bad data, but may be updated for other cases as they are discovered. It will return the location of the bad data files for restructuring
# the array of files that are able to be used for analysis and remove any files that have unusable data.

def BadDatDel(FOLDER,DAT):
    count = 0                     # Counter for bad data files
    os.mkdir(os.path.dirname(os.path.dirname(FOLDER[0]))+'/Bad Data')            # This steps two directories back from the first file in the FOLDER array (all files share the same directory) and then adds a path called Bad Data to the end in order to store all the files that get removed. We don't want to add this folder to the first directory back since that is where all the files that get processed reside. 
    t = os.path.dirname(os.path.dirname(FOLDER[0]))+'/Bad Data'
    
    for i in range(len(DAT)):
        if "Current" in DAT[i].columns:
            x = [DAT[i].loc[:,"Current"].tolist()[idx] < 10**-5 for idx, val in enumerate(DAT[i].loc[:,"Current"].tolist())]   # This creates an array for each set of current data and checks if any value is below 10^-6. If all of them are this low there is a problem with the data. Either the thruster was off when the LP was running or some other failure occured.
            ZC = np.where(DAT[i].loc[:,["Current"]].to_numpy() > 0) # Finding where in file the current crosses the x-axis
            if len(ZC[0]) == 0 or all(x):                                     # Checking if the length of the array is equal to 0. This will mean that there doesn't exist a single point where the data crosses the x-axis. Also checks if all the current values are too low to be a valid LP sweep
                shutil.copy2(FOLDER[i], t)    
                os.remove(FOLDER[i])
                
                count += 1
        else:
            print(f'{count} file(s) have been identified as having incorrect headers.')
            
        
        
        print(f'{count} file(s) have been identified as having bad data and removed from the folder.')
    
        

    return None

# The following function is for reading multiple excel files converted to pandas dataframes. 
# Ensure files are in the .xlsx format otherwise this will throw an error. The purpose of this function
# is to read in the excel data in a form (i.e. arrays) that python can perform operations on or be converted into other data
# types.

def BatchRd(Folder):

    DF=[]

    for i in range(len(Folder)):
        df = pd.read_excel(Folder[i],index_col=False)
        DF.append(df)
    
    return DF

# This function initializes the current and voltages from the raw LP data files. Since it is used multiple times throughout the code a function
# was created for it. This was also created to help better facilitate the iterative part of the analysis if needed.

def VC(data):
    
    DAT = BatchRd(data)
    V = []
    I = []
    
    for i in range(len(DAT)):
        I.append(DAT[i].to_numpy()[:,1])
        V.append(DAT[i].to_numpy()[:,0])
        
    return I, V

# ZeroCross function to take all files to find zero cross of each data file in folder. This helps with calculating the floating potential and provides a
# good reference point for other analytical purposes.

def ZeroCross(data, I):
    DAT = BatchRd(data)
    ZCross = []
    # First we find where each of the current values crosses the x-axis in each of the files, creating a list object of arrays that have all the
    # locations where the current is greater than 0
    for i in range(len(DAT)):
        ZC = np.where(I[i] > 0)
        ZCross.append(ZC[0][0])
    
    return ZCross

# The following function generates a range around the zeros cross point for each of the files for the purpose of later data analysis, specifically, regression fitting to find floating potential 
# since the voltage data will not necessarily always have a value for 0 bias. Currently this function picks 7 data points. This may need updated depending on different bia voltage step sizes.

def ZCAnalysisRng(ZCdat,I,V):
    
    Rng1 = []
    Rng2 = []
    
    for i in range(len(ZCdat)):
        
        rng1 = np.array(I[i])[range(ZCdat[i]-6,ZCdat[i]+5)]
        rng2 = np.array(V[i])[range(ZCdat[i]-6,ZCdat[i]+5)]
        Rng1.append(rng1)
        Rng2.append(rng2)
    
    return Rng1, Rng2

# Function for estimating floating potential for each data file

def FloatingPot(Irange,Vrange):
    
    VFloat = []
    
    # Performs quadratic regression for range around where the zero bias point is. This is probably sufficient in this range of the data even 
    # though the overall shape of an LP curve is not quadratic. In the small region around the zero cross, this should be a good approximation
    # of the function behavior. In some cases even linear regression may be more than adequate, but quadratic should work for cases where it is more linear
    # or if there is more curvature in the region it will better approximate it than using a linear regression method for all cases.
    
    for i in range(len(Irange)):
        P = np.poly1d(np.polyfit(Vrange[i],Irange[i],2))
        if P.deriv(2).coeffs > 0: 
            VF = sorted((P-0).roots)[1]  # Note here we are taking the second root of the solution because the LP curve goes positive to the right. The root of the quadratic going the other direction is in the ion saturation region and will not be a solution because that is not where the LP curve intersects the voltage axis.
        if P.deriv(2).coeffs < 0:
            VF = sorted((P-0).roots)[0]  # This is for the case when the parabola is concave down and the uptrending portion will be the first root in the solution.
        
        VFloat.append(VF)
        
    return VFloat

# The following function is for estimating the plasma potential using numerical methods

def PlasmaPot(data, ZeroX, IE, V):
    # This method solves the second derivative at each data point and then calculates the moving average of the second derivative along
    # the LP trace data. It looks for the spot where the moving average has 3 consecutive negative values (i.e. inflection point is reached)
    # and then uses the voltage associated with that ending point as the plasma potential.
    
    # It also incorporates Lobbia's method of finding the maximum of the first derivative if the second derivative method fails.

    DAT = BatchRd(data)
    
    I_All = []
    V_All = []

    # Need to create ranges from the zero cross point to EOF for each file to perform the analysis on. 
    
    for j in range(len(DAT)):
        i_all = np.array(IE[j])[range(ZeroX[j],len(IE[j]))]
        v_all = np.array(V[j])[range(ZeroX[j],len(V[j]))]

        I_All.append(i_all)
        V_All.append(v_all)

    # Calculating second derivative numerically for instantaneous change of the instantaneous slope for each data file
    # The appropriate moving average window will be up to the analyst, but it is going to be based on a combination of the step size and the
    # range of the data. The moving average window needs to be sized in such a way that it smooths out the data to reflect the general trend.
    # Not so large that changes in the trend at a more local level can't be distinguished and not so small that the data is too "noisy" and
    # no trend can be determined. The smaller the window, the more sensitive the moving average is to changes in the data, the larger the window,
    # the lower the sensitivity to these changes. For the purpose of this type of analysis (trying to find the inflection point of the data)
    # it is suggested to use a window of no more than a few percent of the size of the overall data set. It is left to the analyst's intuition
    # to judge what is an appropriate window size.
    
    # Calculating the second derivative of each data point.
    
    # What would be really cool in future versions of this code was if this script could try different window sizes if
    # the one chosen yields unfavorable results. Would be interesting to see if the window size could be optimized for
    # individual data sets.
    
    D1 = []
    D2 = []
    V_Plasma = []
    MAvg1 = []
    MAvg2 = []
    MaxD = []

    
        
    # We need to smooth out the derivative values. This will be accomplished with a moving average. The size of the moving average
    # is determined by something we will define as a "smoothing factor", selectable by the user of this code. The smoothing factor is expressed
    # in terms of a percent of the full length of the data set. It represents the fraction of the data set that will be used to calculate each moving
    # average. For instance, for a data set of 300 values and a smoothing factor of 1, the size of the moving average is calculated as 300*1/100 = 3.
    # So for this case, the moving average will be calculated using 3 data values. In most cases you won't have an evenly divisible number like this so
    # a ceiling function will be used to determine the total number of data points to use in calculating the moving average. This means that with the 
    # smoothing factor set to 1 and 201 data points, the moving average will still use 3 values.
        
    SmF = 1                 # Set this to whatever the desired smoothing factor is. Currently this is causing issues with the arrays for D. Need to figure out what is going on. Currently window is hard-coded
    
     
    #print(f'Window is {Win}')
    d1 = []
    d2 = []
        
    for i in range(len(DAT)):
        
        Win = 3
        d1 = []
        d2 = []
        
        for j in range(len(V_All[i])-(Win-1)):
            
            # The below is a numerical method for caclulating the first and second derivative using 3 points
            
            D = ((I_All[i][j+2]-I_All[i][j+1])/(V_All[i][j+2]-V_All[i][j+1])-(I_All[i][j+1]-I_All[i][j])/(V_All[i][j+1]-V_All[i][j]))/((V_All[i][j+2]-V_All[i][j])/2)
            d2.append(D)

            d = (I_All[i][j+1]-I_All[i][j])/(V_All[i][j+1]-V_All[i][j])
            
            # The following is to handle cases where the derivative is zero if they exist. This will prevent potential divide-by-zero errors
            # later in the code
            if d == 0:
                d = 10**-9
            d1.append(d)
            
        D1.append(d1)
        D2.append(d2)

        # Calculating moving average of 1st and 2nd derivative 
        # Calculating moving average of 2nd derivative 
        
        
    MaxDev = []
    for i in range(len(DAT)):
        
        Win = 3
        MovAvg1 = []
        MovAvg2 = []
        
        for k in range(len(D1[i])-(Win-1)):
            
            Avg1 = sum(D1[i][k:k+Win])/Win
            Avg2 = sum(D2[i][k:k+Win])/Win
            
            MovAvg1.append(Avg1)
            MovAvg2.append(Avg2)
        
        # Need to take the max of the moving average (equivalent to Dsmooth) and then find in V_All where the index is where this max occurs.
        
        MAvg1.append(MovAvg1)
        MAvg2.append(MovAvg2)
        
        
        # Checking where the 4 consecutive negative values of the second derivative moving average occurs
        count = 0
        it = 0          # Location where the end of the moving average window is

        while count < len(MovAvg2)-Win:
             if MovAvg2[count] < 0 and MovAvg2[count+1] < 0 and MovAvg2[count+2] < 0 and MovAvg2[count+3] < 0 and I_All[i][count] > 0:
                 it = count+(Win-1)
                 break
        
             count += 1
             
        if it == 0:
            while count < len(MovAvg2)-Win:
                if MovAvg2[count] < 0 and MovAvg2[count+1] < 0 and MovAvg2[count+2] < 0 and I_All[i][count] > 0:
                    it = count+(Win-1)
                    break
    
                count += 1
            
        # Considering removing the below. It is possible this is too lenient and can find plasma potential too early in the LP trace. Removing this will
        # cause the script to default to Lobbia's method after failing a stricter criteria.
        '''    
        # The following two statements with the nested while loops are safety nets in case 3 consecutive negative values cannot be found and only 2 or 1 are 
        # identified.
        if it == 0:
            while count < len(MovAvg2)-Win:
                if MovAvg2[count] < 0 and MovAvg2[count+1] < 0:
                    it = count+(Win-1)
                    break

                count += 1

        if it == 0: 
            while count < len(MovAvg2)-Win:
                if MovAvg2[count] < 0:
                    it = count+(Win-1)
                    break

                count += 1
        '''        
        # Bad data error trap. This will return a message to the user if no inflection point in the data is located
        # If this statement is true it also will assign the max derivative calculated earlier to the plasma potential array applying Lobbia's method.
        
        if it == 0:
            
            print('BAD DATA! Plasma potential cannot be found! Second derivative does not go to 0 so an inflection point cannot be identified.')
            v_plasma = V_All[i][np.where(MovAvg1 == max(MovAvg1))[0][0]]
            
        # If the condition is satisfied for the second derivative method then the following code is executed and assigns the potential value calculated 
        # for the location of the consecutive negative moving averages to the plasma potential. 
        
        else: 
            v_plasma = V_All[i][count]
            
        # The 'Offset' variable is added here to account for the fact that the moving average is a lagging indicator. The actual inflection
        # point starts at the beginning of the moving average window. For example, if the moving average window is 3 units, then
        # the first instance of a negative value is 2 units behind the location 'it'.
        
        # The following routine for estimating plasma potential is based on Lobbia's suggestions of using the first derivative. This will
        # be used as a backup method in case the inflection point method fails. This can yield erroneous results if there are any abrupt changes 
        # in the LP sweep data since it uses a max function. In one of the original test files from the work done on the first iteration of the WMGIT
        # thruster this method resulted in a large error in the plasma potential estimation due to a significant discontinuity in the data where the function
        # correctly identified the location of the max first derivative, but this did not correspond to the actual plasma potential. In cases like this,
        # the original method developed for this script should be used as it analyzes the data moving left to right instead of finding a maximum of the derivative 
        # in the entire data set.
          
        
        V_Plasma.append(v_plasma)

    return V_Plasma
    
 # The following routine for estimating plasma potential is based on Lobbia's suggestions of using the first derivative. This will
 # be used as a backup method in case the inflection point method fails. This can yield erroneous results if there are any abrupt changes 
 # in the LP sweep data since it uses a max function. In one of the original test files from the work done on the first iteration of the WMGIT
 # thruster this method resulted in a large error in the plasma potential estimation due to a significant discontinuity in the data where the function
 # correctly identified the location of the max first derivative, but this did not correspond to the actual plasma potential. In cases like this,
 # the original method developed for this script should be used as it analyzes the data moving left to right instead of finding a maximum of the derivative 
 # in the entire data set.
 
def PlasmaPot_v2(data, IE, V):
    
    DAT = BatchRd(data)

    # Identifying the range of current and voltage values beyond the zero cross. The inflection point will be somewhere to the right of the 
    # zero cross point.
    
    # Calculating first derivative for LP electron current by calculation of the instantaneous slope
    D1 = []
    for i in range(len(DAT)):
        d1 = []
        
        for j in range(len(V[i])-1):
            d = (IE[i][j+1]-IE[i][j])/(V[i][j+1]-V[i][j])
            
            # The following is to handle cases where the derivative is zero if they exist. This will prevent potential divide-by-zero errors
            # later in the code
            if d == 0:
                d = 10**-9
            d1.append(d)

        D1.append(np.array(d1)) # Converting type list to numpy array type
        
    # Next we need to smooth out the derivative values. This will be accomplished with a moving average. The size of the moving average
    # is determined by something we will define as a "smoothing factor", selectable by the user of this code. The smoothing factor is expressed
    # in terms of a percent of the full length of the data set. It represents the fraction of the data set that will be used to calculate each moving
    # average. For instance, for a data set of 300 values and a smoothing factor of 1, the size of the moving average is calculated as 300*1/100 = 3.
    # So for this case, the moving average will be calculated using 3 data values. In most cases you won't have an evenly divisible number like this so
    # a ceiling function will be used to determine the total number of data points to use in calculating the moving average. This means that with the 
    # smoothing factor set to 1 and 201 data points, the moving average will still use 3 values.
    
    MAL = []    # Number of points used to calculate moving average
    
    # Calculating for each data file
    for i in range(len(DAT)):
        SmF = 1                 # Set this to whatever the desired smoothing factor is
        mal = math.ceil(len(V[i])*SmF/100)    
        MAL.append(mal)
        
    # Initializing the first set of indices to use for claculating the moving average
    idxs = []
    for i in range(len(MAL)):
    	idxs.append(np.arange(MAL[i]))
    
    Dsmooth = []      # Smoothed derivative
    for i in range(len(DAT)):
        IDX = idxs[i]  
        ma = []
        
        for j in range(len(D1[i])-(MAL[i])):     # Here the derivative set is decremented by the moving average size so the proper length is used for Dsmooth. Using a moving average decreases the total size of the smoothed values.
            ma.append((D1[i][IDX]).mean())       # Calculating the mean of the derivatives within the current moving average window.
            IDX += 1                             # Incrementing moving average window by 1 to calculate the next moving average.
        Dsmooth.append(ma)
    
    V_Plasma = []
    
    # Next the maximum derivative from the smoothed set is found. This is the point where the inflection point occurs in theory.
    for i in range(len(DAT)):
        MaxDev = max(Dsmooth[i])     # Maximum derivative
        
        # Plasma potential is calculated below where the max derivative was located
        V_Plasma.append(V[i][np.where(Dsmooth[i] == MaxDev)[0][0]])
        
    return V_Plasma
# The following function is for calculating the electron current 

def ECurr(VF,data, I, V):
    
    DAT = BatchRd(data)
        
    # Calculating best fit line below floating potential. This will give an equation for the ion saturation current in the form
    # I_isat(Vb) = m_isat*Vb + b_isat where "Vb" is the probe bias, "m" and "b" are the slope and intercept of the line respectively.
    
    # Creating the array of voltage values "x" up until the floating potential. Because the floating potential is calculated and may not match exactly to the
    # data set of voltages we need to find the location in the voltage set where the value is closest to the calculated floating potential. That is the purpose of
    # calculating the V-VF which is the array of differences between the voltage data set and the floating potential. We locate the minimum of this to define our
    # ending point for the voltage "x" values.
       
    I_isat = []
    Ie = []
    slope = []
    
    for i in range(len(DAT)):
        temp = abs(V[i]-VF[i])
        idx = np.where(temp == min(temp))   # Index where value closest to floating potential is located
     
        x = V[i][0:idx[0][0]]       # Voltage values below floating potential
        y = I[i][0:idx[0][0]]       # Current values below floating potential
        model = np.polyfit(x,y,1)   # Returns the slope, "model[0]" and intercept, "model[1]" for the line of best fit in the region
        
        # Creating array of ion saturation current calculations based on the regression model in the previous step for all bias voltages.
        
        iisat = []
        for c in range(len(V[i])):
            predict = np.poly1d(model)
            iisat2 = predict(V[i][c])
            iisat.append(iisat2)
        I_isat.append(iisat)
        slope.append(predict)
        
        # Computing electron current in the form Ie(Vb) = Iprobe(Vb)-I_isat(Vb). This is the correction to the probe current based on the calculated ion saturation current for each probe bias value. 
        # This will be calculated for each probe bias value in the data set so this will be accomplished with a for loop.
        
    for i in range(len(DAT)):  
        ie = []
        for v in range(len(V[i])):
            ie2 = I[i][v] - I_isat[i][v]
            ie.append(ie2)
        Ie.append(ie)
    
    
    return Ie, I_isat, slope
     

# Below is the plasma properties function. This will be used to calculate all the plasma properties of interest 
   
def PlasmaProps(LP_Shape,e,Ap,rp,VF,VP,IE,M,m,Te,ZeroX,data,I,V):
    
    epsilon_0 = 8.8541878188*10**-12   # Permittivity of free space in F/m
    
    DAT = BatchRd(data)

    I_esat = []
    Ne = []
    NE = []
    #print(IE[0])
    for i in range(len(DAT)):
        
        i_esat = IE[i][np.where(V[i] == VP[i])[0][0]]        # Calculating electron saturation current for each file
        ne = i_esat/(e*Ap)*math.sqrt(2*math.pi*m/(e*Te[i]))  # Caclulating electron number density
        I_esat.append(i_esat)
        Ne.append(ne)                      # Value for use in later steps
        NE.append('%.6E' % Decimal(ne))    # Value to return (scientific notation)
        
    print(f'Current is {I_esat}')
    
    # First a single value for the ion saturation current needs to be determined. Since this is calculated as an array
    # of values along a line of best fit. This line also includes positive values, but the ion saturation current is only defined
    # for negative values. The actual data will be used to calculate a point where the ion current stabilizes to a certain point
    # and then average the values to the beginning of the file where the first current value is measured. This should provide a rough estimate 
    # for the ion saturation current since not all data sets will stabilize to a singular value (i.e. not all data sets will show complete saturation).
    
    i_isat = []
    ionsatlocmin = []
    ionsatlocmax = []
    
    S = LP(LP_Shape)             # Instance of the LP class. It is assumed that data files will be processed using one LP shape so this is not
                                 # included in the for loop. If there is reason to have different LP shapes for all the data files this can be
                                 # added as a feature later. For now this will keep it simple so the LP object isn't instantiated every iteration
                                 # of the for loop below.
    
    Vdiff = []
    for i in range(len(ZeroX)):
        
        j = ZeroX[i]
    
        MA = []     # Moving average of current
    
        # The following loop creates an array with a 3 sample moving average
        for j in range(ZeroX[i],1,-1):
        
            M = (I[i][j]+I[i][j-1]+I[i][j-2])/3
            MA.append(M)
            
            vdiff = V[i][j]-V[i][j-1]    # This will be used in calculating the moving average slopes.
            Vdiff.append(vdiff)
            j -= 1
        
            # Next the slope of the moving average is calculated at each step
        k = 0
        D = []
        
        
        for k in range(0,len(MA)-1):
            Diff = (MA[k]-MA[k+1])/Vdiff[k]
            D.append(Diff)    
            k += 1
    
        # Finally a comparison is made between the present and preceding ratios of the slope of the MA. When these are very close to 1 adequate stabilization has been reached. 
        # The following loop will find the locations in the array where these ratios fall below 1.05. These locations will then be used to create a range of values over which the
        # average current is calculated in this stabilization region.
        
        # In future versions of this code, this will be turned into a function where the user may select the stabilization criteria. Here it is hard-coded as 5% and has to be
        # modified manually by the user of this script.
        try:
            w = 0
           
            for w in range(len(D)-1):
                if abs(D[w]/D[w+1]) <= 1.01 and abs(D[w-1]/D[w]) > 1.01:
                   ionsatlocmin.append(w) 
                if w >= 1 and abs(D[w]/D[w+1]) >= 1.01 and abs(D[w-1]/D[w]) < 1.01:
                    ionsatlocmax.append(w)
                    break
                w += 1
        
            #i_isat.append(np.mean(I[i][0:np.where(min(V[i]-VF[i]))[0][0]]))   # Ion saturation current in amps
            i_isat.append(np.mean(I[i][ZeroX[i]-ionsatlocmax[i]:ZeroX[i]-ionsatlocmin[i]]))
        
        except:
            
            i_isat.append(np.mean(I[i][0:ZeroX[i]]))
        
        I_Isat = np.multiply(i_isat,1000)            # Ion saturation current in milliamps
          
    
        # Calculating ion number density assuming thin sheath with the equation ni = -(exp(1/2)*I_isat/eA_s)*sqrt(m_i/e*T_e)
        # After this is completed the debye length will be calculated to determine if this is an appropriate calculation of the ion number
        # density. If thin sheath does not apply then either the transitional sheath or OML model will be applied to solve for
        # ion number density
        
        Ni = []
        for x in range(len(i_isat)):
            ni = -np.exp(1/2)*i_isat[x]/(e*Ap)*np.sqrt(m/(e*Te[x]))     # Ion number density calculated with thin sheath assumption in m^-3
            Ni.append('%.6E' % Decimal(ni))
    
        # Calculating Debye length for each measurement
        
        lambda_D = np.sqrt(np.multiply(epsilon_0,Te)/np.multiply(Ne,e))
        print(f'NE is {Ne}')
        # Calculating the probe radius to debye length ratio to determine correct method for ion number density calculation
        SC = np.multiply(rp,1/lambda_D)
        
        Ni_ud = []                   # Updated ion number density
        
        for i in range(len(SC)):
            
            if SC[i] >= 50:
                ni = Ni[i]
                Ni_ud.append('%.6E' % Decimal(ni))
                
            else: 
                
                # Calculating a and b parameters for ion density correction. See langmuir probe class "Params" method for steps
                
                A = S.Params(SC[i])[0]
                B = S.Params(SC[i])[1]
                
                # To calculate the ion number density we need to calculate the slope between the most negative probe bias and the floating potential
                # and apply an exponent of 1/B to the probe current and calculate the slope within this region.
                
                temp = abs(V[i]-VF[i])
                idx = np.where(temp == min(temp))
                 
                x = V[i][0:idx[0][0]]            # Voltage values below floating potential
                y = I[i][0:idx[0][0]]**(1/B)     # Current values below floating potential with the 1/B exponent for equation 11 in Lobbia
                y_filtered = y[y >= 0]           # Filtering out negative values of current since this can lead to NaN values in certain cases.
                x_filtered = x[y >= 0]           # Resized array of x-values for the filtered y-values
                model = np.polyfit(x_filtered,y_filtered,1)        # Returns the slope, "model[0]" and intercept, "model[1]" for the line of best fit in the region
                
                # Calculating corrected ion density
                ni = 1/(A*Ap)*np.sqrt(2*math.pi*m)*np.exp(-3/2)*Te[i]**(B-0.5)*(-model[0])**B
                Ni_ud.append('%.6E' % Decimal(ni))
                
                
    return NE, Ni_ud, I_Isat, SC, I_esat
            

# This function can be used to plot the data files for visualization. To cut down on processing time, if you have a large folder of
# data make sure to enter how many of the first files in the directory you would like to visualize. This is the "PlotInst" variable
# that is an input to this function which will be implemented at a later time.

def PlotData(data,ZeroX,V_Zerox_rng,I_Zerox_rng,VP,VF):
    
    # The following is for plotting the LP data around the zero cross point and see how well the model fits the data to
    # verify the reliability of the calculation.

    PlotInst = None      # Currently this variable is not defined

    DAT = BatchRd(data)
    V = []
    I = []
    for i in range(len(DAT)):
        V.append(DAT[i].to_numpy()[:,0])
        I.append(DAT[i].to_numpy()[:,1])
    
    
    for j in range(len(DAT)):
        
        polyline = np.linspace(ZeroX[j]+V[j][0]-5,ZeroX[j]+V[j][0]+5,50)
        plt.scatter(V_Zerox_rng[j],I_Zerox_rng[j])
        P = np.poly1d(np.polyfit(V_Zerox_rng[j],I_Zerox_rng[j],2))
        plt.plot(polyline,P(polyline),'r')
        
        plt.title("LP Data At Zero Cross")
        plt.xlabel("Voltage (V)")
        plt.ylabel("Current (A)")
        plt.grid(color='k', linestyle='-', linewidth=0.5)
        plt.show()
        
        plt.scatter(V[j],I[j])
        plt.plot(VP[j],I[j][np.where(V[j]==VP[j])],marker="o", markersize=10, markeredgecolor="red", markerfacecolor="green")
        plt.plot(VF[j],I[j][ZeroX[j]],marker="o", markersize=10, markeredgecolor="black", markerfacecolor="blue")
        
        plt.annotate('Floating Potential', (VF[j],I[j][ZeroX[j]]))
        plt.annotate('Plasma Potential', (VP[j],I[j][np.where(V[j]==VP[j])]))
        plt.title(f'{j} LP Data with Critical Voltage Values')
        plt.xlabel("Voltage (V)")
        plt.ylabel("Current (A)")
        plt.grid(color='k', linestyle='-', linewidth=0.5)
        plt.savefig(f'C:/Users/nnuzz_rpjwflj/OneDrive/Documents/Asif-Nick Business & Thruster Stuff/2025 LP Data Analysis/LP Graph PNGs/LP_Plot_{j}.png')
        plt.show()
        
        
# The following function is used to calculate the electron temperature by using the least squares method from Lobbia after
# transforming the electron current into the natural logarithm of electron current. This function also returns the r-square
# value comparing the fits to the data to determine the quality of the prediction of electron temperature. Anything below 0.9 should
# be ran using the backup plasma potential function.

def ElectronTemp(data, IE, VF, VP, V, M, m):
    
    DAT = BatchRd(data)
   
    Et = []
    R_sq = []
    
    plt.figure()   # Resets the plot on a new figure each time the plotting routine for this function is executed
    
    for i in range(len(DAT)):
        
        # Creating indices for the region between the floating potential and plasma potential
        
        idx0 = np.where(abs(V[i] - VF[i]) == min(abs(V[i] - VF[i])))     # Floating potential index. The absolute value of the difference between the voltage values and the floating potential is used here 
                                                                         # because the calculated floating potential doesn't map to any of the voltage values as it is calculated from a line of best fit.
        idx1 = np.where(V[i] == VP[i])                                   # Plasma potential index
         
        # Need to find where in the region the derivative stops changing substantially. We start by analyzing the derivative
        # dln(I)/dV starting from idx0. Look for when the previous derivative is < 5% of the next. This 5% can be changed depending on preference.
        
        # Note: The below method might benefit from a smoothing algorithm in order to generalize better to other data sets that might have
        # choppy data.
        
        
        k = 0
        C = 0
        
        # This for loop looks at the values from the floating potential to the plasma potential and terminates when it reaches the designated stopping point
        # which is where the beginning of the linear region is identified
        for j in range(idx0[0][0], idx1[0][0]):      
            
            if IE[i][j] > 0 and IE[i][j+1] > 0 and IE[i][j+2] > 0:               # This will only work for non-negative values of IE
                
                d0 = (math.log(IE[i][j+1])-math.log(IE[i][j]))/(V[i][j+1]-V[i][j])
                d1 = (math.log(IE[i][j+2])-math.log(IE[i][j+1]))/(V[i][j+2]-V[i][j+1])
            
                if((d1-d0)/d0) < 0.05:     # If the error between the two successive derivatives is less than 5% increment k by 1
                    k += 1
                if((d1-d0)/d0) > 0.05:     # If the error between the two successive derivatives is greater than 5% reset counter k
                    k = 0
                
                if k == 2:                 # If two consecutive differences in derivatives have been found that are less than 5% exit the for loop 
                    break
        
        # This assigns a value to constant C by increasing the index by 3 past where the two successive low errors were found in the derivative
        # to make sure we are well within the linear region and then subtracts the idx0 value which is where the floating potential so that we
        # can calculate how many indices past the floating potential we need to move to be in the linear region.
        
        C = (j+3)-idx0[0][0]         
        
        # Creating x (voltage) and ln(y) (natural log of current) arrays between the floating and plasma potential
        
        x = V[i][idx0[0][0]+C:idx1[0][0]+2]
        y = np.log(IE[i][idx0[0][0]+C:idx1[0][0]+2])
        
        # Creating x (voltage) and ln(y) (natural log of current) arrays between the floating and plasma potential
        
        try:
            # Generating a linear regression model for this region. 
            model = np.polyfit(x,y,1)
            d = model[0]   # This is just the derivative of the model. Since it is linear, this is just the slope of the linear regression model
            et = 1/d       # The electron temperature is simply the inverse of the slope in this region according to Lobbia  
            Et.append(et)
            
            # Plotting the graphs for the linear regression model for visualization of fit
            p = np.poly1d(model)  # Converting linear regression model coefficients to polynomial function
            lin = np.linspace(x[0],x[-1],len(x))
            plt.grid(color='k', linestyle='-', linewidth=0.5)
            plt.scatter(x,y,label=f'{i}')
            plt.plot(lin, p(lin), '--', color = 'black')
            plt.xlabel("Voltage (V)")
            plt.ylabel("Current [ln(A)]")
            plt.title("Voltage and Natural Log Current Between Floating and Plasma Potentials")
            plt.legend(loc="upper left")
            
            
            # Calculating the correlation coefficients for each of the fits to determine quality of prediction of electron temperature.
            # Correlation coefficients under 0.9 may indicate there is an issue with the plasma potential estimation because the natural logarithm of current
            # as a function of voltage should be mostly linear in this region.
            r_sq = r2_score(y, p(lin))
            R_sq.append(r_sq)
        
        # If above fails default to potential method for calculating electron temperature
        
        except:
            et = (VP[i]-VF[i])/math.log(math.sqrt(M/(2*math.pi*m)))
            Et.append(et)
        
    return Et, R_sq
    
    
 # This function is for calculating the sheath areas for each measurement     
 
def CL_Sheath(rp, Ap, SC, Shape):
    
    
    if Shape == 'cylindrical' or Shape == 'spherical':    
        # Lobbia says to use Vb << Vf  and to ensure (Vp-Vb) >> Te for calculating the sheath thickness. We shall choose an arbitrary
        # Vb = Vp - 10*Te to satisfy the condition (Vp-Vb) >> Te. Assuming this is a factor of 10 allows consistency in calculating the 
        # sheath thickness for all data sets and simplifies our equation for sheath thickness so that it is only dependent on the Debye length
        # as follows.
        
        if isinstance(SC, list):
            xs = [(SC[idx]/rp)**-1*(20**(3/4)*math.sqrt(2)/3) for idx, val in enumerate(SC)]   # Sheath thicknesses in m. Note the PProps function in here. This will have to be ran each time this section is iterated. It will need to accept
                                                                                               # inputs from this section of the code and not file inputs as the original function is written.
            
            # For cylindrical probes, the sheath area is as follows. In later versions the option to do spherical will be added.
            
            As = [Ap*(1+xs[idx]/(d/2)) for idx, val in enumerate(xs)]     # Sheath areas in m^2
        
        else:
            xs = (SC/rp)**-1*(20**(3/4)*math.sqrt(2)/3)   # Sheath thicknesses in m. Note the PProps function in here. This will have to be ran each time this section is iterated. It will need to accept
                                                          # inputs from this section of the code and not file inputs as the original function is written.
            
            # For cylindrical probes, the sheath area is as follows. In later versions the option to do spherical will be added.
            
            As = Ap*(1+xs/(d/2))     # Sheath areas in m^2
            
    if Shape == 'planar':
        As = Ap
    
    return As



# The following function is for writing the results to an excel file.

def WriteExcel(data, FolderLoc, VF, VP, Isat, ET, N):
    
    if Workbook is None or xlwt is None:
        raise ImportError("xlwt is required to use WriteExcel(). Install it with `pip install xlwt`.")

    # Workbook is created 
    wb = Workbook() 
  
    # add_sheet is used to create sheet. 
    sheet1 = wb.add_sheet('Sheet 1') 
    
    # Specifying style 
    style = xlwt.easyxf('font: bold 1') 
    
    # Adding column titles
    sheet1.write(0,0,'File', style)
    sheet1.write(0,1,'Floating Potential (V)', style)
    sheet1.write(0,2,'Plasma Potential (V)', style)
    sheet1.write(0,3,'Ion Saturation Current (mA)', style)
    sheet1.write(0,4,'Electron Temperature (eV)', style)
    sheet1.write(0,5,'Plasma Density (m^-3)', style)
    
    
    # Loading all post-processed data into cells
    for i in range(len(data)):
        sheet1.write(i+1,0,data[i])
        sheet1.write(i+1,1,VF[i])
        sheet1.write(i+1,2,VP[i])
        sheet1.write(i+1,3,Isat[i])
        sheet1.write(i+1,4,ET[i])
        sheet1.write(i+1,5,N[i])
    
    # Make the following save function user input so that files don't get over-written
    wb.save(FolderLoc+'/LP_Data_Post_Processing_Summary_NON-Iterated_Sols_TEST.xls')
    

    
    
#def EEDF_Model(E,Y, e, m, Ap):
    
#    Y = 2/(e**2*Ap)*(2*m*e*E)**(1/2)*
    
    
    
####################################################### END OF FUNCTIONS #################################################################
