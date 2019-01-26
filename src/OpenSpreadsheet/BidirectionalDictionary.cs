namespace OpenSpreadsheet
{
    using System.Collections.Generic;

    /// <summary>
    /// Extends the <see cref="Dictionary{TKey, TValue}"/> collection to support fast reverse lookups.
    /// </summary>
    /// <typeparam name="TKey">The dictionary's key type.</typeparam>
    /// <typeparam name="TValue">The dictionary's value type.</typeparam>
    public sealed class BidirectionalDictionary<TKey, TValue> : Dictionary<TKey, TValue>
    {
        private readonly Dictionary<TValue, TKey> reverseLookup = new Dictionary<TValue, TKey>();

        /// <summary>
        /// Adds a new key-value pair to the collection.
        /// </summary>
        /// <param name="key">The key to be added.</param>
        /// <param name="value">The value to be added.</param>
        public new void Add(TKey key, TValue value)
        {
            base.Add(key, value);
            this.reverseLookup.Add(value, key);
        }

        /// <summary>
        /// Attempts to retrieve a key associated with the provided value.
        /// </summary>
        /// <param name="value">The value being queried.</param>
        /// <param name="key">An output parameter that will contain the found key, if any.</param>
        /// <returns>A value indicating whether a key associated with the provided value was identified.</returns>
        public bool TryGetKey(TValue value, out TKey key)
        {
            if (this.reverseLookup.TryGetValue(value, out key))
            {
                return true;
            }

            key = default;
            return false;
        }
    }
}