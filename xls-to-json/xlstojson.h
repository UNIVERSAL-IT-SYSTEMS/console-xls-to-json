#ifndef __XLSTOJSON_H__
#define __XLSTOJSON_H__

#include <stdlib.h>
#include <stdio.h>
#include <getopt.h>

#include "freexl.h"
#include "libjson.h"
#include "iconv.h"

void xlstojson(char *path);
    
#ifdef _WIN32
#include <string>
void unescape_path(std::string &path);
#endif
#endif