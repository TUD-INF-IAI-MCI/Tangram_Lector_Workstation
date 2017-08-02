using BrailleIO.Dialogs;
using BrailleIO.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace tud.mci.tangram.TangramLector.View
{
    class SceneGraphView
    {
        private Dialog buildSceneGraph()
        {
            string title = "Scene Graph"; // TODO: based on current document/page title?
            string id = "SCENE_GRAPH"; // TODO: multiple ids needed (each for every page)?
            Dialog sceneGraph = new Dialog(title, id);

            #region event registration
            sceneGraph.EntrySelected += new EventHandler<EntryEventArgs>(sceneGraph_EntrySelected);
            sceneGraph.EntryDeselected += new EventHandler<EntryEventArgs>(sceneGraph_EntryDeselected);
            sceneGraph.EntryActivated += new EventHandler<EntryEventArgs>(sceneGraph_EntryActivated);
            sceneGraph.EntryDeactivated += new EventHandler<EntryEventArgs>(sceneGraph_EntryDeactivated);
            sceneGraph.EntryChecked += new EventHandler<EntryEventArgs>(sceneGraph_EntryChecked);
            sceneGraph.EntryUnchecked += new EventHandler<EntryEventArgs>(sceneGraph_EntryUnchecked);
            sceneGraph.SwitchedToParentDialog += new EventHandler<DialogEventArgs>(sceneGraph_SwitchedToParentDialog);
            sceneGraph.SwitchedToChildDialog += new EventHandler<DialogEventArgs>(sceneGraph_SwitchedToChildDialog);
            sceneGraph.EntryAdded += new EventHandler<EntryEventArgs>(sceneGraph_EntryAdded);
            sceneGraph.EntryRemoved += new EventHandler<EntryEventArgs>(sceneGraph_EntryRemoved);
            #endregion

            return sceneGraph;
        }

        #region events

        void sceneGraph_EntryRemoved(object sender, EntryEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_EntryAdded(object sender, EntryEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_SwitchedToChildDialog(object sender, DialogEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_SwitchedToParentDialog(object sender, DialogEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_EntryUnchecked(object sender, EntryEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_EntryChecked(object sender, EntryEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_EntryDeactivated(object sender, EntryEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_EntryActivated(object sender, EntryEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_EntryDeselected(object sender, EntryEventArgs e)
        {
            throw new NotImplementedException();
        }

        void sceneGraph_EntrySelected(object sender, EntryEventArgs e)
        {
            throw new NotImplementedException();
        }

        #endregion

        Dialog _sceneGraph = null;
        /// <summary>
        /// Gets or sets the scene graph.
        /// </summary>
        /// <value>
        /// The scene graph.
        /// </value>
        public Dialog SceneGraph
        {
            get
            {
                if (_sceneGraph == null) _sceneGraph = buildSceneGraph();
                return _sceneGraph;
            }
            set { _sceneGraph = value; }
        }

        /// <summary>
        /// Gets the selected entry.
        /// </summary>
        /// <returns></returns>
        public DialogEntry GetSelectedEntry()
        {
            if (SelectedEntry != null) return (SelectedEntry);
            else return null;
        }

        private DialogEntry _selectedEntry;
        /// <summary>
        /// Gets or sets the current selected entry.
        /// </summary>
        /// <value>
        /// The selected.
        /// </value>
        public DialogEntry SelectedEntry
        {
            get { return _selectedEntry; }
            private set
            {
                if (value == _selectedEntry) return;
                var oldVal = _selectedEntry;
                _selectedEntry = value;
                if (oldVal != null) fire_EntryDeselected(oldVal);
                if (_selectedEntry != null) fire_EntrySelected(_selectedEntry);
            }
        }

        private void fire_EntrySelected(DialogEntry _selectedEntry)
        {
            throw new NotImplementedException();
        }

        private void fire_EntryDeselected(DialogEntry oldVal)
        {
            throw new NotImplementedException();
        }

        #region DialogContentHandling

        /// <summary>
        /// Handles a Tab on a SceneGraph element.
        /// </summary>
        /// <param name="contentAtPosition">The content at position.</param>
        private void handleDialogContent(object contentAtPosition)
        {
            if (contentAtPosition != null && contentAtPosition is DialogEntry)
            {
                DialogEntry dialogEntryAtPosition = (DialogEntry)contentAtPosition;

                // TODO: handle tab
                //if (SceneGraph.GetSelectedEntry() != null)
                //{
                //        handleAudio("msg.audio.Read", ActiveDialog.GetSelectedEntry().Help);
                //}
                //io.RenderDisplay();
            }

        }


        #endregion

    }
}
