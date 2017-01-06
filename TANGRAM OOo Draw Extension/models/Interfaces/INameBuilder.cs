using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace tud.mci.tangram.models.Interfaces
{
    public interface INameBuilder
    {
        /// <summary>
        /// Builds a unique, human-readable name for the object.
        /// </summary>
        /// <returns>A unique name for the object.</returns>
        String BuildName();

        /// <summary>
        /// Rebuilds the name based on the previously build name - e.g. because the name already exists.
        /// </summary>
        /// <returns>A new try for a unique name for the object.</returns>
        String RebuildName();

        /// <summary>
        /// Rebuilds the name based on the previously build name - e.g. because the name already exists.
        /// </summary>
        /// <param name="startIndex">The start index for the next try - e.g. 4 objects of the same type already exists.</param>
        /// <returns>
        /// A new try for a unique name for the object.
        /// </returns>
        String RebuildName(int startIndex);
    }
}