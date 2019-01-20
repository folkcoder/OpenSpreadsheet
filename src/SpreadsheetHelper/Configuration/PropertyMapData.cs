namespace SpreadsheetHelper.Configuration
{
    using System.Reflection;

    using SpreadsheetHelper.Enums;

    /// <summary>
    /// Encapsulates properties associated with an individual property map.
    /// </summary>
    public class PropertyMapData
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyMapData"/> class.
        /// </summary>
        public PropertyMapData() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyMapData"/> class.
        /// </summary>
        /// <param name="propertyInfo">The property data associated with the map.</param>
        public PropertyMapData(PropertyInfo propertyInfo) => this.Property = propertyInfo;

        /// <summary>
        /// Gets or sets the map's column type.
        /// </summary>
        public virtual ColumnType ColumnType { get; set; }

        /// <summary>
        /// Gets or sets a constant value to be used for read and write operations.
        /// </summary>
        public virtual object Constant { get; set; }

        /// <summary>
        /// Gets or sets a constant value to be used for read operations.
        /// </summary>
        public virtual object ConstantRead { get; set; }

        /// <summary>
        /// Gets or sets a constant value to be used for write operations.
        /// </summary>
        public virtual object ConstantWrite { get; set; }

        /// <summary>
        /// Gets or sets a default value to be used for read and write operations when the property value is null.
        /// </summary>
        public virtual object Default { get; set; }

        /// <summary>
        /// Gets or sets a default value to be used for write operations when the property value is null.
        /// </summary>
        public virtual object DefaultRead { get; set; }

        /// <summary>
        /// Gets or sets a default value to be used for write operations when the property value is null.
        /// </summary>
        public virtual object DefaultWrite { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the associated property should be ignored on read and write operations.
        /// </summary>
        public virtual bool Ignore { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the associated property should be ignored on read operations.
        /// </summary>
        public virtual bool IgnoreRead { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the associated property should be ignored on write operations.
        /// </summary>
        public virtual bool IgnoreWrite { get; set; }

        /// <summary>
        /// Sets the one-based column index of the associated property to be used for read and write operations.
        /// </summary>
        public virtual uint Index { get; set; }

        /// <summary>
        /// Sets the one-based column index of the associated property to be used for read operations.
        /// </summary>
        public virtual uint IndexRead { get; set; }

        /// <summary>
        /// Sets the one-based column index of the associated property to be used for write operations.
        /// </summary>
        public virtual uint IndexWrite { get; set; }

        /// <summary>
        /// Sets the column header name to be used for read and write operations.
        /// </summary>
        public virtual string Name { get; set; }

        /// <summary>
        /// Sets the column header name to be used for read operations.
        /// </summary>
        public virtual string NameRead { get; set; }

        /// <summary>
        /// Sets the column header name to be used for write operations.
        /// </summary>
        public virtual string NameWrite { get; set; }

        /// <summary>
        /// Gets the property data associated with the map.
        /// </summary>
        public virtual PropertyInfo Property { get; }

        /// <summary>
        /// Gets or sets the column style.
        /// </summary>
        public virtual ColumnStyle Style { get; set; } = new ColumnStyle();
    }
}