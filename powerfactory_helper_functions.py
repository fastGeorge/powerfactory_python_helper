import powerfactory #type: ignore
import pandas as pd
import numpy as np

class powerfactory_helper_functions:
        
    def __init__(self, app:powerfactory.Application):
        self.app = app

    def clear_events(self):

        int_events = self.app.GetFromStudyCase('IntEvt')
        elm_result = self.app.GetFromStudyCase('ElmRes')

        for evt in int_events.GetContents():
            evt.Delete()

        for evt in elm_result.GetContents():
            evt.Delete()

    def define_shc_event(self, event_name:str, event_object, shc_time, clear_time:float=None):
        int_events = self.app.GetFromStudyCase('*.IntEvt')

        ev_object = int_events.CreateObject('EvtShc', event_name)

        ev_object.p_target = event_object
        ev_object.i_shc = 0 # 3 Phase Short Circuit
        ev_object.time = shc_time

        if type(clear_time) is float:
            ev_clear_obj = int_events.CreateObject('EvtShc', event_name + ' clear shc')
            ev_clear_obj.p_target = event_object
            ev_clear_obj.i_shc = 4 # Clear short circuit
            ev_clear_obj.time = clear_time

    def define_switch_event(self, event_name:str, event_object, ev_time, open_0_close_1:int):
        int_events = self.app.GetFromStudyCase('*.IntEvt')

        ev_object = int_events.CreateObject('EvtSwitch', event_name)

        ev_object.p_target = event_object

        ev_object.time = ev_time

        ev_object.i_switch = open_0_close_1
  
    def export_graph(self, pagename:str, filetype:str):
        """
        Parameters
        ----------
        pagename : str
            Name of the GrpPage that should be exported
        filetype : str
            File type that the export should have, eg. 'png'
        """
        #Set the save path and what to save (Only currently viewed item will be saved)
        ComWr = self.app.GetFromStudyCase('ComWr')
        iopt_rd = filetype
        ComWr.SetAttribute('iopt_rd', iopt_rd)
        ComWr.SetAttribute('iopt_savas', 0) # 0 = Write to path, 1 = Open Save Dialog

        #Get the graphics board
        graphics_board = self.app.GetFromStudyCase('SetDesktop')

        #Get the relevant plot (0 means not creating page in case plot doesn't exist)
        grppage = graphics_board.GetPage(pagename, 0)

        grppage.Show()

        grppage.DoAutoScale()

        ComWr.SetAttribute('f', f'{self.filepath}{pagename}.{iopt_rd}')

        ComWr.Execute()

    def get_simulation_results_as_dataframe(self, elmres = None) -> pd.DataFrame:
        if elmres == None:
            elmres = self.app.GetFromStudyCase('ElmRes')

        elmres.Load()

        d = {}
        for col in range(elmres.GetNumberOfColumns()):
            name = elmres.GetObject(col).loc_name + ";" + elmres.GetVariable(col)

            val_arr = np.zeros(elmres.GetNumberOfRows(), dtype=np.float64)
            if col == 0:
                time_arr = np.zeros(elmres.GetNumberOfRows(), dtype=np.float64)
                for row in range(elmres.GetNumberOfRows()):
                    time_arr[row] = elmres.GetValue(row)[1]
                d["Time"] = pd.Series(time_arr)
            for row in range(elmres.GetNumberOfRows()):
                val_arr[row] = elmres.GetValue(row, col)[1]
            d[name] = pd.Series(val_arr)
        elmres.Flush()

        return pd.DataFrame(d)

    def make_curve(self, pagename, plotname, plot_tuples):
        """
        Parameters
        ----------
        pagename : str
            Name of the GrpPage object where the plot should be created (max 40 characters)
        plotname : str
            Name to be given the plot
        plot_tuples : list
            Should be a list of tuples (DataObject, variable name)
        """
        graphics_board = self.app.GetFromStudyCase('*.SetDesktop')
        #Make a page to plot
        grppage = graphics_board.GetPage(pagename, 1)

        plot = grppage.GetOrInsertCurvePlot(plotname, 1)

        data_series = plot.GetDataSeries()

        data_series.ClearCurves()

        for tuple in plot_tuples:
            data_series.AddCurve(*tuple)

        grppage.DoAutoScale()

        grppage.Show()

    def make_eigenvalue_plot(self, pagename, plotname, type:int):
        """
        Parameters
        ----------
        pagename : str
            Name of the GrpPage object where the plot should be created (max 40 characters)
        plotname : str
            Name to be given the plot
        type : int
            0 eigenvalue plot
            1 mode polar plot
            2 mode bar plot

        Returns: the settings page if additional fixes are required
        """
        graphics_board = self.app.GetFromStudyCase('SetDesktop')
        grppage = graphics_board.GetPage(pagename, 0)

        if grppage != None:
            grppage.RemovePage()
        grppage = graphics_board.GetPage(pagename, 1)

        plot = grppage.GetOrInsertModalAnalysisPlot(plotname, type, create=1)

        res = [x for x in self.app.GetCalcRelevantObjects('*.ElmRes') if 'Modal' in x.loc_name]

        return_ref = None
        if type == 0:
            eigen_plot = plot.GetContents('*.PltEigenvalues')[0]
            eigen_plot.dataTableResultFile = res
            return_ref = eigen_plot
        else:
            mode_plot = plot.GetContents('*.PltEigenmode')[0]
            mode_plot.resultFile = res[0]
            return_ref = mode_plot

        grppage.DoAutoScale()
        grppage.Show()

        return return_ref

    def prepare_rms_simulation(self, int_time_step:float, simulation_time:int):

        init_cond = self.app.GetFromStudyCase('*.ComInc')
        com_sim = self.app.GetFromStudyCase('*.ComSim')
        #Set init conditions
        init_cond.iopt_sim = 'rms' #RMS values
        init_cond.iopt_net = 'sym' #3 phase balanced
        init_cond.iopt_show = 1 #verify initial conditions
        init_cond.dtgrd = int_time_step #set integration time step

        #Set simulation time
        com_sim.tstop = simulation_time

    def set_filepath_for_exports(self, filepath:str):
        self.filepath = filepath

    def set_result_elems(self, pf_object, var_list):
        """
        Parameters
        ----------
        pf_object : DataObject (Elm)
            Object that should get a result from the simulation
        var_list : list
            List of variable names that should get a result from the simulation
        """
        elm_res = self.app.GetFromStudyCase('*.ElmRes')

        elm_res.SetObj(pf_object)
        for var_name in var_list:
            elm_res.AddVariable(pf_object, var_name)


    

        