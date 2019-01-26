namespace OpenSpreadsheet.Configuration
{
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;
    using System.Reflection;

    /// <summary>
    /// Maps class properties to spreadsheet fields.
    /// </summary>
    /// <typeparam name="TClass">The <see cref="Type"/> of class to map.</typeparam>
    public abstract class ClassMap<TClass> where TClass : class
    {
        /// <summary>
        /// Gets a collection of the class map individual property mappings.
        /// </summary>
        public virtual IList<PropertyMap> PropertyMaps { get; } = new List<PropertyMap>();

        /// <summary>
        /// Maps a property to the class map.
        /// </summary>
        /// <typeparam name="TProperty">The property type to be mapped.</typeparam>
        /// <param name="property">The property to be mapped.</param>
        /// <returns>A property map associated with the property.</returns>
        public virtual PropertyMap Map<TProperty>(Expression<Func<TClass, TProperty>> property)
        {
            if (property == null)
            {
                throw new ArgumentNullException(nameof(property));
            }

            if (property.Body is UnaryExpression unaryExp)
            {
                if (unaryExp.Operand is MemberExpression memberExp)
                {
                    var propertyInfo = (PropertyInfo)memberExp.Member;
                    var propertyMap = new PropertyMap(propertyInfo);
                    this.PropertyMaps.Add(propertyMap);
                    return propertyMap;
                }
            }
            else if (property.Body is MemberExpression memberExp)
            {
                var propertyInfo = (PropertyInfo)memberExp.Member;
                var propertyMap = new PropertyMap(propertyInfo);
                this.PropertyMaps.Add(propertyMap);
                return propertyMap;
            }

            throw new ArgumentException($"The expression doesn't indicate a valid property. [ {property} ]");
        }

        /// <summary>
        /// Maps a default property map to the class map; used for constants and values not mapped to a particular class property.
        /// </summary>
        /// <returns>A default property map.</returns>
        public virtual PropertyMap Map()
        {
            var propertyMap = new PropertyMap();
            this.PropertyMaps.Add(propertyMap);
            return propertyMap;
        }
    }
}