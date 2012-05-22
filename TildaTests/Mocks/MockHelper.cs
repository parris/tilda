using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace TildaTests.Mocks {
    class MockHelper {

        /**
         * Makes a deep copy for some object where its class can be [Serializable]
         * @param T a generic object
         * @return T a deep copy of that object
         */
        public static T DeepClone<T>(T obj) {
            using(var ms = new MemoryStream()) {
                var formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                ms.Position = 0;

                return (T)formatter.Deserialize(ms);
            }
        }
    }
}
