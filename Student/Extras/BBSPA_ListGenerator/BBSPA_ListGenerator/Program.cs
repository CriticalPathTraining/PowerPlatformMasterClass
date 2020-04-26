using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BBSPA_ListGenerator {
  class Program {
    static void Main() {

      // create Customers list and add a dozen customer items
      // SharePointListFactory.CreateCustomersList(12, 12);


      SharePointListFactory.CreateAllLists();

    }
  }
}
