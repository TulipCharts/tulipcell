/*
 * Tulip Cell
 * https://tulipcell.org/
 * Copyright (c) 2010-2016 Tulip Charts LLC
 * Lewis Van Winkle (LV@tulipcharts.org)
 *
 * This file is part of Tulip Cell.
 *
 * Tulip Cell is free software: you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as published by the
 * Free Software Foundation, either version 3 of the License, or (at your
 * option) any later version.
 *
 * Tulip Cell is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 * FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public License
 * for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with Tulip Cell.  If not, see <http://www.gnu.org/licenses/>.
 *
 */



#include "indicators.h"


/* This is a generic interface useful for interacting with
 * other programming languages and tools.
 */


__stdcall const char *GetVersion() {
    /* Returns Tulip Indicators version. */
    return TI_VERSION;
}

__stdcall int GetIndicator(const char *name) {
    /* Returns the indicator index for a given name. */
    /* Returns -1 if not found. */
    const ti_indicator_info *info = ti_find_indicator(name);
    if (!info) return -1;
    return info - ti_indicators;
}


__stdcall int GetIndicatorCount() {
    /* Simply returns the number of indicators available. */
    return TI_INDICATOR_COUNT;
}


__stdcall const char *GetName(int index) {
    /* This will return the indicator's name. */
    /* Returns null for not found. */

    if (index <= 0) return 0;
    if (index >= TI_INDICATOR_COUNT) return 0;
    return ti_indicators[index].name;
}


__stdcall const char *GetFullName(int index) {
    /* This will return the indicator's full name. */
    /* Returns null for not found. */

    if (index <= 0) return 0;
    if (index >= TI_INDICATOR_COUNT) return 0;
    return ti_indicators[index].full_name;
}


__stdcall int GetInputCount(int index) {
    /* This will return the number of inputs the indicator expects. */
    /* Returns -1 if not found. */
    if (index <= 0) return -1;
    if (index >= TI_INDICATOR_COUNT) return -1;
    return ti_indicators[index].inputs;
}


__stdcall int GetOptionCount(int index) {
    /* This will return the number of options the indicator expects. */
    /* Returns -1 if not found. */
    if (index <= 0) return -1;
    if (index >= TI_INDICATOR_COUNT) return -1;
    return ti_indicators[index].options;
}


__stdcall int GetOutputCount(int index) {
    /* This will return the number of outputs the indicator expects. */
    /* Returns -1 if not found. */
    if (index <= 0) return -1;
    if (index >= TI_INDICATOR_COUNT) return -1;
    return ti_indicators[index].outputs;
}


__stdcall const char *GetInputName(int index, int input) {
    /* This will return the indicator's input name. */
    /* 0 is the first index. */
    /* Returns null for not found or out-of-bounds. */

    if (index <= 0) return 0;
    if (index >= TI_INDICATOR_COUNT) return 0;
    if (input <= 0) return 0;
    if (input >= ti_indicators[index].inputs) return 0;
    return ti_indicators[index].input_names[input];
}


__stdcall const char *GetOptionName(int index, int option) {
    /* This will return the indicator's option name. */
    /* 0 is the first index. */
    /* Returns null for not found or out-of-bounds. */

    if (index <= 0) return 0;
    if (index >= TI_INDICATOR_COUNT) return 0;
    if (option <= 0) return 0;
    if (option >= ti_indicators[index].options) return 0;
    return ti_indicators[index].option_names[option];
}


__stdcall const char *GetOutputName(int index, int output) {
    /* This will return the indicator's output name. */
    /* 0 is the first index. */
    /* Returns null for not found or out-of-bounds. */

    if (index <= 0) return 0;
    if (index >= TI_INDICATOR_COUNT) return 0;
    if (output <= 0) return 0;
    if (output >= ti_indicators[index].outputs) return 0;
    return ti_indicators[index].output_names[output];
}


__stdcall int GetStart(int index, TI_REAL const *options) {
    /* This will return how much shorter the output for an indicator is
     * than the input, for a given set of options. */
    /* Returns -1 for not found. */

    if (index <= 0) return -1;
    if (index >= TI_INDICATOR_COUNT) return -1;
    return ti_indicators[index].start(options);
}


__stdcall int Call(int index, int size, TI_REAL const *inputs, TI_REAL const *options, TI_REAL *outputs) {
    /* This will run an indicator on data. */
    /* Returns -1 for not found. */
    /* Returns TI_OKAY for okay (0). */
    /* Returns TI_XXX for error. */

    /* This expects the input and output arrays to be continuous. */
    /* For example, if size=5 and we expect two inputs, then inputs should be
     * 10 elements long. The first 5 elements being input 0, and the second 5
     * elements being input 1. Outputs works the same way. */
    /* This will also offset and zero-fill the tops of the output arrays. */
    /* That is so the outputs go to the end of the output arrays. This is very
     * different from how the C interface works. */

    if (index <= 0) return -1;
    if (index >= TI_INDICATOR_COUNT) return -1;


    const double *in[TI_MAXINDPARAMS] = {0};
    double *out[TI_MAXINDPARAMS] = {0};

    const ti_indicator_info *info = &ti_indicators[index];
    const int start = info->start(options);

    int i, j;
    for (i = 0; i < info->inputs; ++i) {
        in[i] = inputs + i * size;
    }
    for (i = 0; i < info->outputs; ++i) {
        out[i] = outputs + i * size + start;
        for (j = 0; j < start; ++j) {
            out[i][j] = 0.0;
        }
    }

    return ti_indicators[index].indicator(size, in, options, out);
}

