#include "pch.h"            //precompiler header.
#include <windows.h>        //excel processing needs.
#include <algorithm>        //for sort.
#include <comutil.h>        //COM.
#pragma comment(lib, "comsuppw.lib")    //linker use, necessary.
#include <vector>


//create struct that will be used for storing each row from the table.
struct RowData
{
    std::vector<_variant_t> cells;  //set vector datatype to that of what Excel will provide: variant.
    //long originalIndex;
};

extern "C" __declspec(dllexport)    //disables changing name of function so VBA can use this one freely.

//create the actual Sort function that will be referenced by VBA.
//params:
//tableVariant : VARIANT, table data passed by Excel to DLL.
//selectedColIndex : long, column index of chosen column to sort.
//errorCode : long, for debugging purposes.
void __stdcall SortTableByColumn(VARIANT* tableVariant, long* selectedColIndex, long* errorCode)
{
    //if table or column index does not exist, return error code.
    *errorCode = 0;
    if (!tableVariant || !selectedColIndex) { 
        *errorCode = 1; 
        return; 
    }

    //store column index retrieved from VBA to a local variable to be used by the DLL.
    long colIndex = *selectedColIndex;

    //check if Variant passed from Excel is by reference or actual value.
    //if passed by reference, store value being referenced by variant to varArray.
    //if not, varArray will keep value from tableVariant.
    //this is used to confirm that variant is acceptable for processing.
    VARIANT* varArray = tableVariant;
    if (tableVariant->vt & VT_BYREF) varArray = tableVariant->pvarVal;

    //if variant does not have safeArray with elements variant, it is not acceptable for processing.
    if (!((varArray->vt & VT_ARRAY) && (varArray->vt & VT_VARIANT))) { 
        *errorCode = 2; return; 
    }

    //store variant to a more acessible variable pointer.
    SAFEARRAY* sa = varArray->parray;

    //l means lower, u means upper.
    //get the lower and upper bounds of the row and column of the variant.
    //confirm that all 4 values are present, otherwise a variant with unknown lower or upper bound
    //is not a legit table variant.
    LONG l_Boundrow, u_Boundrow, l_Boundcol, u_Boundcol;
    if (FAILED(SafeArrayGetLBound(sa, 1, &l_Boundrow)) || FAILED(SafeArrayGetUBound(sa, 1, &u_Boundrow)) ||
        FAILED(SafeArrayGetLBound(sa, 2, &l_Boundcol)) || FAILED(SafeArrayGetUBound(sa, 2, &u_Boundcol)))
    {
        *errorCode = 3; 
        return;
    }

    //get the number or rows and columns.
    //note: since Excel passes a column index that is 1-based, 
    // in order to accurately confirm number of rows and columns, we add 1.
    //e.g upper bound of row is 4, lower bound of row is 1.
    //if we count by hand, we can easily say the number of rows is 4.
    //to compute for this value using the upper and lower bound, we do:
    //upper bound - lower bound + 1.
    LONG rows = u_Boundrow - l_Boundrow + 1;
    LONG cols = u_Boundcol - l_Boundcol + 1;

    //if selected column index is < 1, it is invalid.
    //if selected column index is greater that the total number of columns, it is invalid.
    if (colIndex < 1 || colIndex > cols) { 
        *errorCode = 7; 
        return; 
    }

    try
    {
        //create vector with struct RowData with size of total number of rows.
        std::vector<RowData> rowVector(rows);
        _variant_t val;

        // store table variant into vector.
        for (LONG i = 0; i < rows; ++i)
        {
            //set each row's cell size to number of columns.
            //the cells here represent the number of values inside a row.
            //e.g a row consists of number, name, date and position : 4.
            rowVector[i].cells.resize(cols);
            for (LONG j = 0; j < cols; ++j)
            {
                //create index array variable that would mark each cell's position.
                //we add i and j to the lower bounds of row and col, respectively.
                //this is to convert DLL indeces to match safearray indices.
                LONG idx[2] = { i + l_Boundrow, j + l_Boundcol };

                //get value from safearray with specified index, and store to val.
                SafeArrayGetElement(sa, idx, &val);

                //populate rowVector cells with acquired value from safearray.
                //this will then be used for sorting.
                rowVector[i].cells[j] = val;
            }
        }

        // sort.
        std::sort(rowVector.begin(), rowVector.end(),
            [colIndex](const RowData& a, const RowData& b)
            {
                //create to variables for comparing.
                //for the index, we -1 because column index from excel is 1-based.
                //our rowVector is 0-based.
                _variant_t va = a.cells[colIndex - 1];
                _variant_t vb = b.cells[colIndex - 1];        

                //VT_I4 - number.
                if (va.vt == VT_I4 && vb.vt == VT_I4)
                    return va.lVal < vb.lVal;

                //VT_DATE - date from excel.
                if (va.vt == VT_R8 || va.vt == VT_I4 || va.vt == VT_DATE)
                    return va.date < vb.date;

                // string.
                _bstr_t sa(va);
                _bstr_t sb(vb);
                return wcscmp((wchar_t*)sa, (wchar_t*)sb) < 0;
            });

        // store back to safearray, which will be passed back to VBA.
        for (LONG i = 0; i < rows; ++i)
        {
            for (LONG j = 0; j < cols; ++j)
            {
                LONG idx[2] = { i + l_Boundrow, j + l_Boundcol };
                SafeArrayPutElement(sa, idx, &rowVector[i].cells[j]);
            }
        }
    }
    catch (...)
    {
        *errorCode = 5;
    }
}


