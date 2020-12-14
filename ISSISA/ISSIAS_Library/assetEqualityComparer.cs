using System.Collections.Generic;

namespace ISSIAS_Library
{
    public class assetEqualityComparer : IEqualityComparer<asset>
    {
        #region IEqualityComparer<asset> Members
        public bool Equals(asset x, asset y)
        {
            return x.asset_number.Equals(y.asset_number);
        }

        public int GetHashCode(asset obj)
        {
            unchecked
            {
                var hash = 17;
                //same here, if you only want to get a hashcode on a, remove the line with b
                hash = hash * 23 + obj.GetHashCode();
                hash = hash * 23 + obj.GetHashCode();
                return hash;
            }
        }
        #endregion
    }
}
