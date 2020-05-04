

import com.jacob.com.SafeArray;
import com.jacob.com.Variant;

public class COM_Basic_Commands_utils
{
    private final static String DIMENSION_ERROR = "Wrong Dimension in SafeArray!";

    public static Variant[][] get2DArrayFromSafeArray(SafeArray safeArray) throws Exception
    {
        if (isTwoDimensional(safeArray))
        {
            int size1 = getFirstDimensionSize(safeArray);
            int size2 = getSecondDimensionSize(safeArray);
            Variant[][] array = new Variant[size1][size2];
            for (int i = 0; i < size1; i++)
            {
                for (int j = 0; j < size2; j++)
                {
                    array[i][j] = safeArray.getVariant(i, j);
                }
            }
            return array;
        }
        throw new Exception(DIMENSION_ERROR);
    }

    public static int[][] get2DIntArrayFromSafeArray(SafeArray safeArray) throws Exception
    {
        if (isTwoDimensional(safeArray))
        {
            int size1 = getFirstDimensionSize(safeArray);
            int size2 = getSecondDimensionSize(safeArray);
            int[][] array = new int[size1][size2];
            for (int i = 0; i < size1; i++)
            {
                for (int j = 0; j < size2; j++)
                {
                    array[i][j] = safeArray.getInt(i, j);
                }
            }
            return array;
        }
        throw new Exception(DIMENSION_ERROR);
    }


    public static String[][] get2DStringArrayFromSafeArray(SafeArray safeArray) throws Exception
    {
        if (isTwoDimensional(safeArray))
        {
            int size1 = getFirstDimensionSize(safeArray);
            int size2 = getSecondDimensionSize(safeArray);
            String[][] array = new String[size1][size2];
            for (int i = 0; i < size1; i++)
            {
                for (int j = 0; j < size2; j++)
                {
                    array[i][j] = safeArray.getString(i, j);
                }
            }
            return array;
        }
        throw new Exception(DIMENSION_ERROR);
    }

    public static double[][] get2DDoubleArrayFromSafeArray(SafeArray safeArray) throws Exception
    {
        if (isTwoDimensional(safeArray))
        {
            int size1 = getFirstDimensionSize(safeArray);
            int size2 = getSecondDimensionSize(safeArray);
            double[][] array = new double[size1][size2];
            for (int i = 0; i < size1; i++)
            {
                for (int j = 0; j < size2; j++)
                {
                    array[i][j] = safeArray.getDouble(i, j);
                }
            }
            return array;
        }
        throw new Exception(DIMENSION_ERROR);
    }

    private static int getFirstDimensionSize(SafeArray safeArray)
    {
        return safeArray.getUBound(1) + 1;
    }

    private static int getSecondDimensionSize(SafeArray safeArray)
    {
        return safeArray.getUBound(2) + 1;
    }

    private static boolean isTwoDimensional(SafeArray safeArray)
    {
        return safeArray.getNumDim() == 2;
    }

}
