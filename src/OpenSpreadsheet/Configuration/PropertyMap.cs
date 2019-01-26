namespace OpenSpreadsheet.Configuration
{
    using System.Reflection;

    using Enums;

    /// <summary>
    /// Mapping info between a class property and a spreadsheet column.
    /// </summary>
    public class PropertyMap
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyMap"/> class.
        /// </summary>
        public PropertyMap() => this.PropertyData = new PropertyMapData();

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyMap"/> class.
        /// </summary>
        /// <param name="propertyInfo">The property data associated with the map.</param>
        public PropertyMap(PropertyInfo propertyInfo) => this.PropertyData = new PropertyMapData(propertyInfo);

        /// <summary>
        /// Gets the data associated with the property map.
        /// </summary>
        public PropertyMapData PropertyData { get; }

        /// <summary>
        /// Sets the column type for both read and write operations.
        /// </summary>
        /// <param name="columnType">The column type to be applied.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap ColumnType(ColumnType columnType)
        {
            this.PropertyData.ColumnType = columnType;
            return this;
        }

        /// <summary>
        /// Sets a constant value to be used for both read and write operations.
        /// </summary>
        /// <param name="value">The constant value to be used. When reading, the value must be the same type as the mapped property.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap Constant(object value)
        {
            this.PropertyData.Constant = value;
            this.PropertyData.ConstantRead = value;
            this.PropertyData.ConstantWrite = value;

            return this;
        }

        /// <summary>
        /// Sets a constant value to be used for read operations.
        /// </summary>
        /// <param name="value">The constant value to be used. The value must be of the same type as the mapped property.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap ConstantRead(object value)
        {
            this.PropertyData.ConstantRead = value;
            return this;
        }

        /// <summary>
        /// Sets a constant value to be used for write operations.
        /// </summary>
        /// <param name="value">The constant value to be used.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap ConstantWrite(object value)
        {
            this.PropertyData.ConstantWrite = value;
            return this;
        }

        /// <summary>
        /// Sets a default value to be used for read and write operations when the property value is null.
        /// </summary>
        /// <param name="value">The default value to be used. When reading, the value must be the same type as the mapped property.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap Default(object value)
        {
            this.PropertyData.Default = value;
            this.PropertyData.DefaultRead = value;
            this.PropertyData.DefaultWrite = value;

            return this;
        }

        /// <summary>
        /// Sets a default value to be used for read operations when the property value is null.
        /// </summary>
        /// <param name="value">The default value to be used. The value must be the same type as the mapped property.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap DefaultRead(object value)
        {
            this.PropertyData.DefaultRead = value;
            return this;
        }

        /// <summary>
        /// Sets a default value to be used for write operations when the property value is null.
        /// </summary>
        /// <param name="value">The default value to be used.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap DefaultWrite(object value)
        {
            this.PropertyData.DefaultWrite = value;
            return this;
        }

        /// <summary>
        /// Sets a value indicating whether the associated property should be ignored on read and write operations.
        /// </summary>
        /// <param name="value">A value indicating whether the associated property should be ignored.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap Ignore(bool value)
        {
            this.PropertyData.Ignore = value;
            this.PropertyData.IgnoreRead = value;
            this.PropertyData.IgnoreWrite = value;

            return this;
        }

        /// <summary>
        /// Sets a value indicating whether the associated property should be ignored on read operations.
        /// </summary>
        /// <param name="value">A value indicating whether the associated property should be ignored.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap IgnoreRead(bool value)
        {
            this.PropertyData.IgnoreRead = value;
            return this;
        }

        /// <summary>
        /// Sets a value indicating whether the associated property should be ignored on write operations.
        /// </summary>
        /// <param name="value">A value indicating whether the associated property should be ignored.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap IgnoreWrite(bool value)
        {
            this.PropertyData.IgnoreWrite = value;
            return this;
        }

        /// <summary>
        /// Sets the one-based column index of the associated property to be used for read and write operations.
        /// </summary>
        /// <param name="index">The property's column index.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap Index(uint index)
        {
            this.PropertyData.Index = index;
            this.PropertyData.IndexRead = index;
            this.PropertyData.IndexWrite = index;

            return this;
        }

        /// <summary>
        /// Sets the one-based column index of the associated property to be used for read operations.
        /// </summary>
        /// <param name="index">The property's column index.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap IndexRead(uint index)
        {
            this.PropertyData.IndexRead = index;
            return this;
        }

        /// <summary>
        /// Sets the one-based column index of the associated property to be used for write operations.
        /// </summary>
        /// <param name="index">The property's column index.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap IndexWrite(uint index)
        {
            this.PropertyData.IndexWrite = index;
            return this;
        }

        /// <summary>
        /// Sets the column header name to be used for read and write operations.
        /// </summary>
        /// <param name="name">The property's column header name.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap Name(string name)
        {
            this.PropertyData.Name = name;
            this.PropertyData.NameRead = name;
            this.PropertyData.NameWrite = name;

            return this;
        }

        /// <summary>
        /// Sets the column header name to be used for read operations.
        /// </summary>
        /// <param name="name">The property's column header name.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap NameRead(string name)
        {
            this.PropertyData.NameRead = name;
            return this;
        }

        /// <summary>
        /// Sets the column header name to be used for read operations.
        /// </summary>
        /// <param name="name">The property's column header name.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap NameWrite(string name)
        {
            this.PropertyData.NameWrite = name;
            return this;
        }

        /// <summary>
        /// Sets the column style.
        /// </summary>
        /// <param name="columnStyle">The style to be applied.</param>
        /// <returns>The changed PropertyMap.</returns>
        public virtual PropertyMap Style(ColumnStyle columnStyle)
        {
            this.PropertyData.Style = columnStyle;
            return this;
        }
    }
}