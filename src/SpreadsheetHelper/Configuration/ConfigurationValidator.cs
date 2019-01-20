namespace SpreadsheetHelper.Configuration
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    /// <summary>
    /// Provides methods to validate configuration properties.
    /// </summary>
    public class ConfigurationValidator
    {
        private readonly IEnumerable<PropertyMap> propertyMaps;

        /// <summary>
        /// Gets a collection of validation errors.
        /// </summary>
        public List<Exception> Errors { get; } = new List<Exception>();

        /// <summary>
        /// Gets a value indicating whether the validator has errors.
        /// </summary>
        public bool HasErrors => this.Errors.Count > 0;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationValidator"/> class.
        /// </summary>
        /// <param name="propertyMaps">The property maps to be validated.</param>
        public ConfigurationValidator(IEnumerable<PropertyMap> propertyMaps) => this.propertyMaps = propertyMaps;

        /// <summary>
        /// Validates the property maps.
        /// </summary>
        public void Validate()
        {
            this.Errors.Clear();

            // read
            this.ValidateIndexesAreUnique(this.propertyMaps.Where(x => !x.PropertyData.IgnoreRead && x.PropertyData.IndexRead > 0).Select(x => x.PropertyData.IndexRead), ConfigurationType.Read);
            this.ValidateReadProperties(this.propertyMaps.Where(x => !x.PropertyData.IgnoreRead && x.PropertyData.Property != null));
            foreach (var map in this.propertyMaps.Where(x => !x.PropertyData.IgnoreRead))
            {
                this.ValidateConstant(map.PropertyData.ConstantRead, map);
                this.ValidateDefault(map.PropertyData.DefaultRead, map);
                this.ValidateIndexHasValue(map.PropertyData.IndexRead, map, ConfigurationType.Read);
                this.ValidateIndexWithinExcelMaxRange(map.PropertyData.IndexRead, map, ConfigurationType.Read);
            }

            // write
            this.ValidateIndexesAreUnique(this.propertyMaps.Where(x => !x.PropertyData.IgnoreWrite && x.PropertyData.IndexWrite > 0).Select(x => x.PropertyData.IndexWrite), ConfigurationType.Write);
            foreach (var map in this.propertyMaps.Where(x => !x.PropertyData.IgnoreWrite))
            {
                this.ValidateIndexHasValue(map.PropertyData.IndexWrite, map, ConfigurationType.Write);
                this.ValidateIndexWithinExcelMaxRange(map.PropertyData.IndexWrite, map, ConfigurationType.Write);
                this.ValidateHeaderNameWithinExcelMaxLength(map.PropertyData.NameWrite, map);
            }

            this.Errors.OrderBy(x => x.Message);
        }

        private void ValidateConstant(object constantValue, PropertyMap map)
        {
            if (constantValue == null)
            {
                return;
            }

            // constant is not mapped to a particular property
            if (map.PropertyData.Property == null)
            {
                return;
            }

            var propertyType = Nullable.GetUnderlyingType(map.PropertyData.Property.PropertyType) ?? map.PropertyData.Property.PropertyType;
            if (constantValue.GetType() != propertyType)
            {
                this.Errors.Add(new ArgumentException($"Constant of type '{constantValue.GetType().FullName}' does not match member of type '{map.PropertyData.Property.PropertyType.Name}'."));
            }
        }

        private void ValidateDefault(object defaultValue, PropertyMap map)
        {
            if (defaultValue == null)
            {
                return;
            }

            var propertyType = Nullable.GetUnderlyingType(map.PropertyData.Property.PropertyType) ?? map.PropertyData.Property.PropertyType;
            if (defaultValue.GetType() != propertyType)
            {
                this.Errors.Add(new ArgumentException($"Default of type '{defaultValue.GetType().FullName}' does not match member of type '{map.PropertyData.Property.PropertyType.Name}'."));
            }
        }

        private void ValidateHeaderNameWithinExcelMaxLength(string name, PropertyMap map)
        {
            const int maxHeaderLength = 255;

            string headerName = name ?? map.PropertyData.Property.Name;
            if (headerName.Length > maxHeaderLength)
            {
                this.Errors.Add(new ArgumentException($"Property '{map.PropertyData.Property.Name}' has a header name '{headerName}' that is longer than the maximum length allowed by Excel ({maxHeaderLength.ToString()})."));
            }
        }

        private void ValidateIndexesAreUnique(IEnumerable<uint> indexes, ConfigurationType configurationType)
        {
            var uniqueIndexes = new HashSet<uint>();
            foreach (var index in indexes)
            {
                if (uniqueIndexes.Contains(index))
                {
                    this.Errors.Add(new ArgumentException($"Column index {index} is defined for multiple {configurationType.ToString().ToLower()} properties."));
                }
                else
                {
                    uniqueIndexes.Add(index);
                }
            }
        }

        private void ValidateIndexHasValue(uint index, PropertyMap map, ConfigurationType configurationType)
        {
            if (index == 0)
            {
                this.Errors.Add(new ArgumentException($"{configurationType.ToString()} property '{map.PropertyData.Property.Name}' does not have a defined column index. The Excel column collection begins at index 1."));
            }
        }

        private void ValidateIndexWithinExcelMaxRange(uint index, PropertyMap map, ConfigurationType configurationType)
        {
            const int maxColumnIndex = 16384;

            if (index > maxColumnIndex)
            {
                this.Errors.Add(new ArgumentException($"{configurationType.ToString()} property '{map.PropertyData.Property.Name}' has a defined column index '{index.ToString()}' that is greater than the maximum number of columns allowed by Excel ({maxColumnIndex.ToString()})."));
            }
        }

        private void ValidateReadProperties(IEnumerable<PropertyMap> propertyMaps)
        {
            var props = new HashSet<PropertyInfo>();
            foreach (var map in propertyMaps)
            {
                if (props.Contains(map.PropertyData.Property))
                {
                    this.Errors.Add(new ArgumentException($"Read property '{map.PropertyData.Property.Name}' is mapped to more than one column."));
                }
                else
                {
                    props.Add(map.PropertyData.Property);
                }
            }
        }

        internal enum ConfigurationType
        {
            Read,
            Write
        }
    }
}