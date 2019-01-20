namespace SpreadsheetHelper.Enums
{
    using System;

    /// <summary>
    /// Encapsulates valid cell border placement values.
    /// </summary>
    [Flags]
    public enum BorderPlacement : uint
    {
        /// <summary>
        /// No borders.
        /// </summary>
        None = 0,

        /// <summary>
        /// Border on the left side of the cell.
        /// </summary>
        Left = 1 << 0,

        /// <summary>
        /// Border on the right side of the cell.
        /// </summary>
        Right = 1 << 1,

        /// <summary>
        /// Border on the top side of the cell.
        /// </summary>
        Top = 1 << 2,

        /// <summary>
        /// Border on the bottom side of the cell.
        /// </summary>
        Bottom = 1 << 3,

        /// <summary>
        /// Border on the entire outside of the cell.
        /// </summary>
        Outside = Left | Right | Top | Bottom,

        /// <summary>
        /// Border from the bottom left to the top right of the cell.
        /// </summary>
        DiagonalUp = 1 << 4,

        /// <summary>
        /// Border from the bottom right to the top left of the cell.
        /// </summary>
        DiagonalDown = 1 << 5,

        /// <summary>
        /// All possible borders.
        /// </summary>
        All = Outside | DiagonalUp | DiagonalDown
    }
}