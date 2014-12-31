#                                               -*- Autoconf -*-
# Process this file with autoconf to produce a configure script.

AC_PREREQ([2.69])
AC_INIT([xlstojson], [1.0], [keisuke.miyako@4d.com])
AM_INIT_AUTOMAKE([xlstojson], [1.0])
AC_CONFIG_SRCDIR([xlstojson.cc])
# AC_CONFIG_HEADERS([config.h])

# Checks for programs.
AC_PROG_CXX
AC_PROG_CC

# Checks for libraries.
# FIXME: Replace `main' with a function in `-lfreexl':
AC_CHECK_LIB([freexl], [freexl_close])
# FIXME: Replace `main' with a function in `-ljson':
AC_CHECK_LIB([json], [json_free])

# Checks for header files.
AC_CHECK_HEADERS([stdlib.h])

# Checks for typedefs, structures, and compiler characteristics.

# Checks for library functions.

AC_CONFIG_FILES([Makefile])
AC_OUTPUT
