/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package workdaycalculator;

/**
 *
 * @author Mike
 */
public class Tools {
    public int convert(String p_pos) {
        if (p_pos == null || p_pos.isEmpty()) {
            return 0;
        }
        int res = 0;
        int len = p_pos.length();
        for (int i = 0; i < len; i++) {
            res *= 26;
            res += p_pos.charAt(i) - 64;
        }
        return res;

    }

}
