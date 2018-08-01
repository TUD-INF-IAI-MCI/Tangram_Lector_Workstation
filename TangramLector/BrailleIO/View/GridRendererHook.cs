using BrailleIO.Interface;
using System;
using System.Collections.Generic;
using System.Linq;

namespace tud.mci.tangram.TangramLector.BrailleIO.View
{
    /// <summary>
    /// A renderer hook providing an overlaying grid pattern. 
    /// </summary>
    /// <seealso cref="BrailleIO.Interface.IBailleIORendererHook" />
    /// <seealso cref="BrailleIO.Interface.IActivatable" />
    class GridRendererHook : IBailleIORendererHook, IActivatable
    {
        #region Members

        /// <summary>
        /// The configuration key prefix
        /// </summary>
        protected static readonly string CONFIG_KEY = typeof(GridRendererHook).Name + "_";

        /// <summary>
        /// Sets the horizontal vertical grid size = the distance between two horizontal or 
        /// vertical grid lines.
        /// </summary>
        /// <value>
        /// The size of the grid.
        /// </value>
        public virtual int GridSize { set { GridSizeHorizontal = GridSizeVertical = value; } }
        /// <summary>
        /// Sets the lines pattern for horizontal and vertical lines of the grid.
        /// The pattern is a defining array alternating between length of raised pins and gaps.
        /// E.g. [1, 2] means, a 1 pin line is followed by a 2 pin gap a.s.o. .
        /// </summary>
        /// <value>
        /// The grid line pattern.
        /// </value>
        public virtual int[] GridLinePattern { set { GridHorizontalLinesPattern = GridVerticalLinesPattern = value; } }
        /// <summary>
        /// Sets the horizontal and vertical offset for the grid start.
        /// </summary>
        /// <value>
        /// The grid offset.
        /// </value>
        public virtual int GridOffset { set { GridOffsetHorizontal = GridOffsetVertical = value; } }

        /// <summary>
        /// Gets or sets the horizontal grid size = the distance between two horizontal grid lines.
        /// </summary>
        /// <value>
        /// The horizontal grid size.
        /// </value>
        public virtual int GridSizeHorizontal { get; set; }

        int _goh = 0;
        /// <summary>
        /// Gets or sets the horizontal offset for the grid start.
        /// </summary>
        /// <value>
        /// The horizontal grid offset.
        /// </value>
        public virtual int GridOffsetHorizontal
        {
            get { return _goh; }
            set
            {
                if (value < 1)
                {
                    _goh = (GridSizeHorizontal + ((value) % GridSizeHorizontal)) % GridSizeHorizontal;
                }
                else if (value >= GridSizeHorizontal)
                {
                    _goh = (value) % GridSizeHorizontal;
                }
                else _goh = (value);
            }
        }

        /// <summary>
        /// Gets or sets the vertical grid size = the distance between two vertical grid lines.
        /// </summary>
        /// <value>
        /// The vertical grid size.
        /// </value>
        public virtual int GridSizeVertical { get; set; }

        int _gov = 0;
        /// <summary>
        /// Gets or sets the vertical offset for the grid start.
        /// </summary>
        /// <value>
        /// The vertical grid offset.
        /// </value>
        public virtual int GridOffsetVertical
        {
            get { return _gov; }
            set
            {
                if (value < 1)
                {
                    _gov = (GridSizeVertical + ((value) % GridSizeVertical)) % GridSizeVertical;
                }
                else if (value >= GridSizeVertical)
                {
                    _gov = (value) % GridSizeVertical;
                }
                else _gov = (value);
            }
        }

        int[] _gph = new int[2] { 1, 0 };
        /// <summary>
        /// Gets or sets the lines pattern for horizontal lines of the grid.
        /// The pattern is a defining array alternating between length of raised pins and gaps.
        /// E.g. [1, 2] means, a 1 pin line is followed by a 2 pin gap a.s.o. .
        /// </summary>
        /// <value>
        /// The pattern for horizontal lines.
        /// </value>
        public virtual int[] GridHorizontalLinesPattern
        {
            get
            {
                if (_gph == null) _gph = new int[0];
                return _gph;
            }
            set
            {
                if (value == null) _gph = new int[0];
                else if (value.Length % 2 == 0) { _gph = value; }
                else
                {
                    _gph = new int[value.Length + 1];
                    value.CopyTo(_gph, 0);
                }
            }
        }

        int[] _gpv = new int[2] { 1, 0 };
        /// <summary>
        /// Gets or sets the lines pattern for vertical lines of the grid.
        /// The pattern is a defining array alternating between length of raised pins and gaps.
        /// E.g. [1, 2] means, a 1 pin line is followed by a 2 pin gap a.s.o. .
        /// </summary>
        /// <value>
        /// The pattern for vertical lines.
        /// </value>
        public virtual int[] GridVerticalLinesPattern
        {
            get
            {
                if (_gpv == null) _gpv = new int[0];
                return _gpv;
            }
            set
            {
                if (value == null) _gpv = new int[0];
                else if (value.Length % 2 == 0) { _gpv = value; }
                else
                {
                    _gpv = new int[value.Length + 1];
                    value.CopyTo(_gpv, 0);
                }
            }
        }

        /// <summary>
        /// Determines if there should be a gap of one lowered pin between 
        /// grid points and raised pins of the underlying picture.
        /// </summary>
        public bool Gap = true;

        #region IActivatable

        bool _active = false;
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="IActivatable" /> is active or not.
        /// </summary>
        /// <value>
        ///   <c>true</c> if active; otherwise, <c>false</c>.
        /// </value>
        public bool Active
        {
            get { return _active && Mode != GridMode.None; }
            set { _active = value; }
        }

        #endregion

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="GridRendererHook"/> class.
        /// </summary>
        public GridRendererHook()
        {
            Active = true;

            loadParameterFromConfig();

            GridSize = 10;
            GridOffset = 0;
           // Gap = false;

            //GridLinePattern = new int[5] { 2, 2, 3, 3, 4 };
            //GridVerticalLinesPattern = new int[2] { 3, 2 };

            // GridLinePattern = new int[2] { 1, 4 };
            //GridLinePattern = new int[2] { 1, 1 };


            //GridVerticalLinesPattern = null;
            //GridVerticalLinesPattern = null;

            //GridHorizontalLinesPattern = new int[2] { 1, 1 };
            //GridVerticalLinesPattern= new int[2] { 1, 1 };



            ///************* 5 FULL ******************/
            //GridSize = 5;
            //GridLinePattern = new int[2] { 100, 0 };


            ///************* 10 FULL ******************/
            //GridSize = 10;
            //GridLinePattern = new int[2] { 100, 0 };


            ///************* 10 DASHED ******************/
            ///* FIXME: moving pattern while panning (oÓ) */
            //GridSize = 10;
            //GridLinePattern = new int[2] { 2, 2 };


            ///************* 10 DOTED ******************/
            //GridSize = 10;
            //GridLinePattern = new int[2] { 1, 1 };


            ///************* 9 DOTED WIDE ******************/
            //GridSize = 9;
            //GridLinePattern = new int[2] { 1, 2 };


            ///************* 10 DOT ******************/
            //GridSize = 10;
            //GridLinePattern = new int[2] { 1, 9 };


            ///************* 5 DOT ******************/
            //GridSize = 5;
            //GridLinePattern = new int[2] { 1, 4 };


            ///************* 10 small Cross ******************/
            //GridSize = 10;
            //GridLinePattern = new int[4] { 2, 7, 1, 0 };


            ///************* 10 Aiming Cross ******************/
            //GridSize = 10;
            //GridLinePattern = new int[6] { 0, 1, 2, 5, 2, 0 };

        }

        #endregion

        #region IBailleIORendererHook

        /// <summary>
        /// This hook function is called by an IBrailleIOHookableRenderer before he starts his rendering.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        /// <param name="additionalParams">Additional parameters.</param>
        public void PreRenderHook(ref IViewBoxModel view, ref object content, params object[] additionalParams)
        {
            return;
        }


        int _lastXOffset = 0;
        int _lastYOffset = 0;

        /// <summary>
        /// This hook function is called by an IBrailleIOHookableRenderer after he has done his rendering before returning the result.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        /// <param name="result">The result matrix, may be manipulated. Addressed in [y, x] notation.</param>
        /// <param name="additionalParams">Additional parameters.</param>
        public void PostRenderHook(IViewBoxModel view, object content, ref bool[,] result, params object[] additionalParams)
        {
            if (Active && view != null && result != null)
            {
                int n = result.GetLength(1);    // amount of columns = width
                int m = result.GetLength(0);    // amount of rows = height

                if (n * m > 0)
                {
                    // check for offset changes
                    if (view is IPannable)
                    {
                        if (_lastXOffset != ((IPannable)view).GetXOffset())
                            GridOffsetHorizontal = _lastXOffset = ((IPannable)view).GetXOffset();
                        if (_lastYOffset != ((IPannable)view).GetYOffset())
                            GridOffsetVertical = _lastYOffset = ((IPannable)view).GetYOffset();
                    }
                    paintGrid(ref result, n, m);
                }
            }
        }

        #endregion

        #region Grid Paint


        private void paintGrid(ref bool[,] result, int n, int m)
        {
            if (Gap) resetGridPointLookup(); // init grid point lookup

            #region draw horizontal lines
            if (GridHorizontalLinesPattern.Length > 1)
            {
                // calculate pattern position based of grid offset
                int o = 0;
                int p1 = 0; // incrementP(0, GridOffsetVertical, GridHorizontalLinesPattern, out o);
                for (int i = GridOffsetVertical - 1; i < m; i += GridSizeVertical) // row
                {
                    int p = p1;

                    if (i < 0) continue;

                    //// if offset leads to start with a gap
                    //if (p % 2 == 1)
                    //{
                    //    i += o;
                    //    p++; o = 0;
                    //}

                    bool ignore = false;
                    bool prevcheck = false;

                    for (int j = GridOffsetHorizontal - GridSizeHorizontal - 1; j < n; j++) // col // FIXME: start position of the pattern
                    {
                        int l = GridHorizontalLinesPattern[p % GridHorizontalLinesPattern.Length];
                        p++;
                        int g = GridHorizontalLinesPattern[p % GridHorizontalLinesPattern.Length];
                        p++;

                        // paint line 
                        for (int z = o; z < l; z++)
                        {
                            int jz = j + z;
                            if (jz > -1 && jz < n)
                            {
                                #region  horizontal Gap handling
                                if (Gap)
                                {

                                    // check 8 neighbors
                                    if (
                                        // 1st: current col
                                        result[i, jz] ||                    // current point is risen
                                        (i > 0 && result[i - 1, jz]) ||     // top point
                                        (i < m - 1 && result[i + 1, jz])    // bottom point                                                
                                    ||
                                        // 2nd: next col
                                        ((jz + 1 < n) &&                            // next col
                                            (result[i, jz + 1] ||                   // next point
                                            (i > 0 && result[i - 1, jz + 1]) ||     // next top right
                                            (i < m - 1 && result[i + 1, jz + 1])))  // next bottom right point
                                        )
                                    {
                                        ignore = true;
                                        z++;
                                        continue;
                                    }

                                    // 3rd: check previous col - only necessary if a gap was before
                                    else if (prevcheck)
                                    {
                                        prevcheck = false;
                                        if ((jz - 1 > -1) &&                        // prev col
                                            (result[i, jz - 1] ||                   // prev point
                                            (i > 0 && result[i - 1, jz - 1]) ||     // prev top right
                                            (i < m - 1 && result[i + 1, jz - 1])))  // prev bottom right point
                                        {
                                            z++;
                                            continue;
                                        }
                                    }

                                    // if the previous point identifies that this one should not be set
                                    if (ignore)
                                    {
                                        ignore = false;
                                        continue;
                                    }

                                    // add to grid point lookup 
                                    add2GridPointLookup(i, jz);
                                }
                                #endregion

                                // set the grid point
                                if (jz > 0) result[i, jz] = true;
                            }
                            else break;
                        }

                        // jump gap size
                        j += Math.Max(0, l + g - 1);
                        if (g > 0)
                        {
                            ignore = false;
                            prevcheck = true;
                        }
                    }
                }
            }
            #endregion

            #region draw vertical lines
            if (GridVerticalLinesPattern.Length > 1)
            {
                // calculate pattern position based of grid offset
                int o = 0;
                int p1 = 0; // incrementP(0, GridOffsetHorizontal, GridVerticalLinesPattern, out o);

                for (int j = GridOffsetHorizontal - 1; j < n; j += GridSizeHorizontal) // col
                {
                    if (j < 0) continue;

                    int p = p1;
                    //// if offset leads to start with a gap
                    //if (p % 2 == 1)
                    //{
                    //    j += o;
                    //    p++; o = 0;
                    //}

                    bool ignore = false;
                    bool prevcheck = false;

                    for (int i = GridOffsetVertical - GridSizeVertical - 1; i < m; i++) // row // FIXME: Start position of the pattern
                    {
                        int l = GridVerticalLinesPattern[p % GridVerticalLinesPattern.Length];
                        p++;
                        int g = GridVerticalLinesPattern[p % GridVerticalLinesPattern.Length];
                        p++;

                        // paint line 
                        for (int z = o; z < l; z++)
                        {
                            int iz = i + z;
                            if (iz > -1 && iz < m)
                            {
                                #region  vertical Gap handling
                                if (Gap)
                                {
                                    // check 8 neighbors
                                    if (
                                        // 1st: current row
                                        (result[iz, j] && !GridPointLookupContains(iz, j)) ||                    // current point is risen
                                        (j > 0 && (result[iz, j - 1] && !GridPointLookupContains(iz, j - 1))) ||     // left point
                                        (j < n - 1 && (result[iz, j + 1] && !GridPointLookupContains(iz, j + 1)))    // right point                                                
                                    ||
                                        // 2nd: next row
                                        ((iz + 1 < m) &&                            // next col
                                            ((result[iz + 1, j] && !GridPointLookupContains(iz + 1, j)) ||                   // bottom point
                                            (j > 0 && (result[iz + 1, j - 1] && !GridPointLookupContains(iz + 1, j - 1))) ||     // bottom left point
                                            (j < n - 1 && (result[iz + 1, j + 1] && !GridPointLookupContains(iz + 1, j + 1)))))  // bottom right point
                                        )
                                    {
                                        ignore = true;
                                        z++;
                                        continue;
                                    }

                                    // 3rd: check previous row - only necessary if a gap was before
                                    else if (prevcheck)
                                    {
                                        prevcheck = false;
                                        if ((iz - 1 > -1) &&                        // prev row
                                            ((result[iz - 1, j] && !GridPointLookupContains(iz - 1, j)) ||                   // top point
                                            (j > 0 && (result[iz - 1, j - 1] && !GridPointLookupContains(iz - 1, j - 1))) ||     // top left 
                                            (j < n - 1 && (result[iz - 1, j + 1] && !GridPointLookupContains(iz - 1, j + 1)))))  // top right point
                                        {
                                            z++;
                                            continue;
                                        }
                                    }

                                    // if the previous point identifies that this one should not be set
                                    if (ignore)
                                    {
                                        ignore = false;
                                        continue;
                                    }
                                }
                                #endregion

                                // set the grid point
                                if (iz > 0) result[iz, j] = true;
                            }
                            else break;
                        }

                        // jump gap size
                        i += Math.Max(0, l + g - 1);
                        if (g > 0)
                        {
                            ignore = false;
                            prevcheck = true;
                        }
                    }
                }
            }
            #endregion
        }

        private static int incrementP(int p, int j, int[] pattern, out int rest)
        {
            rest = 0;
            return 0;

            // FIXME: make this work

            if (pattern != null && pattern.Length > 0)
            {
                int pos = 0;
                while (j > pos)
                {
                    pos += pattern[p % pattern.Length];
                    p++;
                }
                rest = Math.Abs(j - pos);
            }
            return p;
        }

        #endregion

        #region grid point lookup

        List<int> mList = new List<int>();
        Dictionary<int, List<int>> nList = new Dictionary<int, List<int>>();

        readonly object lookupLock = new object();

        void resetGridPointLookup()
        {
            lock (lookupLock)
            {
                mList = new List<int>();
                nList = new Dictionary<int, List<int>>();
            }
        }

        void add2GridPointLookup(int i, int j)
        {
            lock (lookupLock)
            {
                mList.Add(i);
                if (nList.ContainsKey(i))
                {
                    nList[i].Add(j);
                }
                else
                {
                    nList.Add(i, new List<int>(1) { j });
                }
            }
        }

        bool GridPointLookupContains(int i, int j)
        {
            lock (lookupLock)
            {
                var result = mList.Contains(i) && nList[i].Contains(j);
                return result;
            }
        }

        #endregion

        #region Mode handling

        /// <summary>
        /// Gets or sets the mode which kind of grid should be rendered.
        /// </summary>
        /// <value>
        /// The mode.
        /// </value>
        public GridMode Mode { get; set; }

        static readonly int maxMode = Enum.GetValues(typeof(GridMode)).Cast<int>().Max() + 1;
        /// <summary>
        /// Rotate through the grid modes
        /// </summary>
        /// <returns></returns>
        public GridMode NextMode()
        {
            var nextMode = (((int)Mode) + 1) % maxMode;
            var newMode = (GridMode)nextMode;
            return Mode = newMode;
        }

        #endregion

        #region configuration

        void loadParameterFromConfig()
        {
            try
            {
                var config = ConfigurationLoader.GetGlobalConfiguration();

                _laodGapFromConfig(config);
                _laodActiveFromConfig(config);

                _laodGridSizeFromConfig(config);
                _laodGridOffsetFromConfig(config);
                _laodGridLinePatternFromConfig(config);

                _laodGridHorizontalLinesPatternFromConfig(config);
                _laodGridSizeHorizontalFromConfig(config);
                _laodGridOffsetHorizontalFromConfig(config);

                _laodGridVerticalLinesPatternFromConfig(config);
                _laodGridSizeVerticalFromConfig(config);
                _laodGridOffsetVerticalFromConfig(config);
            }
            catch (Exception e)
            {
                Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR]\tfailure while loading configuration keys:", e);
            }
        }

        private void _laodGridLinePatternFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
            if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridLinePattern"))
            {
                string pattern = ConfigurationLoader.GetValueFromConfig(CONFIG_KEY + "GridLinePattern", config);
                if (!String.IsNullOrWhiteSpace(pattern))
                {
                    GridLinePattern = pattern.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
                }
            }
        }
        private void _laodGridHorizontalLinesPatternFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
            if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridHorizontalLinesPattern"))
            {
                string pattern = ConfigurationLoader.GetValueFromConfig(CONFIG_KEY + "GridHorizontalLinesPattern", config);
                if (!String.IsNullOrWhiteSpace(pattern))
                {
                    GridHorizontalLinesPattern = pattern.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
                }
            }
        }
        private void _laodGridVerticalLinesPatternFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
            if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridVerticalLinesPattern"))
            {
                string pattern = ConfigurationLoader.GetValueFromConfig(CONFIG_KEY + "GridVerticalLinesPattern", config);
                if (!String.IsNullOrWhiteSpace(pattern))
                {
                    GridVerticalLinesPattern = pattern.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
                }
            }
        }

        private void _laodGridOffsetFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
             if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridOffset"))
                 GridOffset = ConfigurationLoader.GetValueFromConfig<int>(CONFIG_KEY + "GridOffset", config);
        }

        private void _laodGridSizeFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
            if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridSize"))
            GridSize = ConfigurationLoader.GetValueFromConfig<int>(CONFIG_KEY + "GridSize", config);
        }

        private void _laodGapFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
            if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "Gap"))
            Gap = ConfigurationLoader.GetValueFromConfig<bool>(CONFIG_KEY + "Gap", config);
        }

        private void _laodActiveFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
             if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "Active"))
                 Active = ConfigurationLoader.GetValueFromConfig<bool>(CONFIG_KEY + "Active", config);
        }

        private void _laodGridSizeVerticalFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
             if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridSizeVertical"))
                 GridSizeVertical = ConfigurationLoader.GetValueFromConfig<int>(CONFIG_KEY + "GridSizeVertical", config);
        }

        private void _laodGridSizeHorizontalFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
             if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridSizeHorizontal"))
                 GridSizeHorizontal = ConfigurationLoader.GetValueFromConfig<int>(CONFIG_KEY + "GridSizeHorizontal", config);
        }

        private void _laodGridOffsetVerticalFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
             if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridOffsetVertical"))
                 GridOffsetVertical = ConfigurationLoader.GetValueFromConfig<int>(CONFIG_KEY + "GridOffsetVertical", config);
        }

        private void _laodGridOffsetHorizontalFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
             if (ConfigurationLoader.ConfigContainsKey(CONFIG_KEY + "GridOffsetHorizontal"))
                 GridOffsetHorizontal = ConfigurationLoader.GetValueFromConfig<int>(CONFIG_KEY + "GridOffsetHorizontal", config);
        }

        #endregion

    }


    /// <summary>
    /// The available modes for this 
    /// </summary>
    public enum GridMode : int
    {
        /// <summary>
        /// No grid is rendered
        /// </summary>
        None,
        /// <summary>
        /// Only the grid id rendered
        /// </summary>
        Grid
    }
        
}
