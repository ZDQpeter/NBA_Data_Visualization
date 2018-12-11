import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import os, sys
from function_main_Plotting_Rank import function_main_Plotting_Rank
from function_main_Plotting_Pt_Gain import function_main_Plotting_Pt_Gain
from function_main_Plotting_Pt_Loss import function_main_Plotting_Pt_Loss
from function_main_Plotting_Pt_Diff import function_main_Plotting_Pt_Diff
from function_main_Plotting_Whole_Win_Pctg import function_main_Plotting_Whole_Win_Pctg
from function_main_Plotting_Home_Win_Pctg import function_main_Plotting_Home_Win_Pctg
from function_main_Plotting_Road_Win_Pctg import function_main_Plotting_Road_Win_Pctg
from function_main_Plotting_Division_Win_Pctg import function_main_Plotting_Division_Win_Pctg
from function_main_Plotting_Conference_Win_Pctg import function_main_Plotting_Conference_Win_Pctg
from function_main_Plotting_NBA_Win_Pctg import function_main_Plotting_NBA_Win_Pctg

def function_main_Plotting():

    # ========================================
    # Define parameters
    Rank = 1
    Team = 2
    Wins = 3
    Loss = 4
    Win_Pctg = 5
    Win_Diff = 6
    Home_Result = 7
    Road_Result = 8
    Div_Result = 9
    Conf_Result = 10
    Pt_Gain = 11
    Pt_Loss = 12
    Pt_Diff = 13
    WL_Streak = 14

    # ==========================================================================
    selected_Column = Rank
    function_main_Plotting_Rank(selected_Column)

    # ==========================================================================
    selected_Column = Pt_Gain
    function_main_Plotting_Pt_Gain(selected_Column)

    # ==========================================================================
    selected_Column = Pt_Loss
    function_main_Plotting_Pt_Loss(selected_Column)

    # ==========================================================================
    selected_Column = Pt_Diff
    function_main_Plotting_Pt_Diff(selected_Column)

    # ==========================================================================
    selected_Column = Win_Pctg
    function_main_Plotting_Whole_Win_Pctg(selected_Column)

    # ==========================================================================
    selected_Column = Home_Result
    function_main_Plotting_Home_Win_Pctg(selected_Column)

    # ==========================================================================
    selected_Column = Road_Result
    function_main_Plotting_Road_Win_Pctg(selected_Column)

    # ==========================================================================
    selected_Column = Div_Result
    function_main_Plotting_Division_Win_Pctg(selected_Column)

    # ==========================================================================
    selected_Column = Conf_Result
    function_main_Plotting_Conference_Win_Pctg(selected_Column)

    # ==========================================================================
    selected_Column = Win_Pctg
    function_main_Plotting_NBA_Win_Pctg(selected_Column)

    # =============================================
    plt.show()

    return
