using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BLL.Models
{
    public enum Marker
	{
        Ignore,
        ReminderZUS, //tylko załączniki dotyczące płatności składek ZUS
        ReminderZUS_PIT //tylko załączniki dotyczące płatności składek ZUS_PIT
    }
}
