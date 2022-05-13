﻿// copyright(c) 2016 Stephen Workman (workman.stephen@gmail.com)

using System;

namespace LinqToExcelModern.Logging {
   public interface ILogManagerFactory {
      ILogProvider GetLogger(Type name);
      ILogProvider GetLogger(String name);
   }
}
