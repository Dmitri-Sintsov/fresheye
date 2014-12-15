/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.openoffice.uniyar.fresheye;

import java.util.HashSet;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.beans.XPropertySet;
import com.sun.star.text.XText;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextCursor;

/**
 *
 * @author Administrator
 */
public class FreshEyeProcessor extends TextProcessor {

    private class WordStruc {

        XTextRange para;
        String word;
        int pos;
        Boolean hilited;

        public WordStruc( XTextRange para, String word, int pos, Boolean hilited ) {
            this.para = para;
            this.word = word;
            this.pos = pos;
            this.hilited = hilited;
        }

        public void hilite() {
            int CharColor = 0;
            int CharBackColor = 0;
            if ( !this.hilited ) {
                XText xText = this.para.getText();
                XTextCursor xTextCursor = xText.createTextCursor();
                XPropertySet xCursorProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xTextCursor );
                xTextCursor.gotoRange( this.para, false );
                xTextCursor.collapseToStart();
                xTextCursor.goRight( (short) this.pos, false );
                xTextCursor.goRight( (short) this.word.length(), true );
                try {
                    CharBackColor = Integer.parseInt( (xCursorProps.getPropertyValue( "CharBackColor" )).toString() );
                    CharColor = Integer.parseInt( (xCursorProps.getPropertyValue( "CharColor" )).toString() );
                    if ( CharBackColor == -1 ) {
                        CharBackColor = 0x00FFFF00;
                    } else {
                        CharBackColor = (0xFFFFFFFF - CharBackColor) & 0x00FFFFFF;
                    }
                    if ( CharColor != -1 ) {
                        if ( Math.abs( (long)CharColor - (long)CharBackColor ) < 10000 ) {
                            CharColor = (0xFFFFFFFF - CharColor) & 0x00FFFFFF;
                        }
                    }
                    try {
                        xCursorProps.setPropertyValue( "CharBackColor", CharBackColor );
                        xCursorProps.setPropertyValue( "CharColor", CharColor );
                    } catch( Exception e2 ) {
                        System.err.println( "Could not set CharBackColor property value of xTextCursor : " + e2 );
                        e2.printStackTrace( System.out );
                        System.exit( 0 );
                    }
                } catch( Exception e ) {
                    System.err.println( "Could not get CharBackColor property value of xTextCursor : " + e );
                    e.printStackTrace( System.out );
                    System.exit( 0 );
                }
                this.hilited = true;
            }
        }

    }

    // fresh eye definitions
    private int Threshold_val;
    private static final String REGEXP_LETTER = "[\u0410-\u042F\u0430-\u044F\u0401\u0451]";
    private static final int CONTEXT_LENGTH = 9; // "united" MAXWIDTH / width

    double twosigmasqr = 2 * Math.pow( CONTEXT_LENGTH * 4, 2 );
    int[] psychlen = new int[ CONTEXT_LENGTH ]; // "former" razd (psychological length of word separators)
    WordStruc[] wordlist = new WordStruc[ CONTEXT_LENGTH ]; // "former" w (LIFO of words)

    Boolean firstpara = true;
    Boolean spaces = false;

    HashSet<String> voc_except = new HashSet<String>(); // exceptions vocabulary - "former" voc
    int inf_letters[][] = {  // quantity of information in letters: relative, average = 1000
        // by itself - in the beginning of a word
        {  802,    959 },  /* а */
        { 1232,   1129 },  /* б */
        {  944,    859 },  /* в */
        { 1253,   1193 },  /* г */
        { 1064,    951 },  /* д */
        {  759,   1232 },  /* е */
        {  759,   1232 },  /* ё */
        { 1432,   1432 },  /* ж */
        { 1193,    993 },  /* з */
        {  802,    767 },  /* и */
        { 1329,   1993 },  /* й */
        { 1032,    929 },  /* к */
        {  967,   1276 },  /* л */
        { 1053,    944 },  /* м */
        {  848,    711 },  /* н */
        {  695,    853 },  /* о */
        { 1088,    454 },  /* п */
        {  929,   1115 },  /* р */
        {  895,    793 },  /* с */
        {  848,   1002 },  /* т */
        { 1115,   1129 },  /* у */
        { 1793,   1022 },  /* ф */
        { 1259,   1329 },  /* х */  /* {0} manually decreased! was 1359 */
        { 1593,   1393 },  /* ц */
        { 1276,   1212 },  /* ч */
        { 1476,   1012 },  /* ш */
        { 1676,   1676 },  /* щ */
        { 1993,   3986 },  /* ъ */
        { 1193,   3986 },  /* ы */
        { 1253,   3986 },  /* ь */
        { 1676,   1232 },  /* э */
        { 1476,   1793 },  /* ю */
        { 1159,    967 }   /* я */
    };
    int sim_ch[][] = { // letters' similarity map
      /* а б в г д е ё ж з и й к л м н о п р с т у ф х ц ч ш щ ъ ы ь э ю я */
        {9,0,0,0,0,1,1,0,0,1,0,0,0,0,0,2,0,0,0,0,1,0,0,0,0,0,0,0,1,0,1,1,2}, /* а */
        {0,9,1,0,0,0,0,0,0,0,0,0,0,0,0,0,3,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0}, /* б */
        {0,1,9,1,0,0,0,0,0,0,0,0,1,1,1,0,1,0,0,0,1,3,0,0,0,0,0,0,0,0,0,0,0}, /* в */
        {0,0,1,9,0,0,0,3,0,0,0,3,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0}, /* г */
        {0,0,0,0,9,0,0,0,1,0,0,0,0,0,0,0,0,0,1,3,0,0,0,1,0,0,0,0,0,0,0,0,0}, /* д */
        {1,0,0,0,0,9,9,0,0,2,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,1,0,2,1,1}, /* е */
        {1,0,0,0,0,9,9,0,0,2,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,1,0,2,1,1}, /* ё */
        {0,0,0,3,0,0,0,9,3,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,3,3,0,0,0,0,0,0}, /* ж */
        {0,0,0,0,1,0,0,3,9,0,0,0,0,0,0,0,0,0,3,1,0,0,0,3,1,1,1,0,0,0,0,0,0}, /* з */
        {1,0,0,0,0,2,2,0,0,9,3,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,2,0,1,1,1}, /* и */
        {0,0,0,0,0,0,0,0,0,2,9,0,1,1,1,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0}, /* й */
        {0,0,0,3,0,0,0,0,0,0,0,9,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0}, /* к */
        {0,0,1,0,0,0,0,0,0,0,1,0,9,1,1,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0}, /* л */
        {0,0,1,0,0,0,0,0,0,0,1,0,1,9,3,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0}, /* м */
        {0,0,1,0,0,0,0,0,0,0,1,0,1,3,9,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0}, /* н */
        {2,0,0,0,0,1,1,0,0,1,0,0,0,0,0,9,0,0,0,0,1,0,0,0,0,0,0,0,1,0,1,1,1}, /* о */
        {0,3,1,0,0,0,0,0,0,0,0,0,0,0,0,0,9,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0}, /* п */
        {0,0,0,0,0,0,0,0,0,0,1,0,1,1,1,0,0,9,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0}, /* р */
        {0,0,0,0,1,0,0,0,3,0,0,0,0,0,0,0,0,0,9,1,0,0,0,3,1,0,0,0,0,0,0,0,0}, /* с */
        {0,0,0,0,3,0,0,0,1,0,0,0,0,0,0,0,0,0,1,9,0,0,0,1,1,0,0,0,0,0,0,0,0}, /* т */
        {1,0,1,0,0,1,1,0,0,1,0,0,0,0,0,1,0,0,0,0,9,0,0,0,0,0,0,0,1,0,1,2,1}, /* у */
        {0,1,3,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,9,0,0,0,0,0,0,0,0,0,0,0}, /* ф */
        {0,0,0,1,0,0,0,0,0,0,0,1,0,0,0,0,0,1,0,0,0,0,9,0,1,0,0,0,0,0,0,0,0}, /* х */
        {0,0,0,0,1,0,0,0,3,0,0,0,0,0,0,0,0,0,3,1,0,0,0,9,0,0,0,0,0,0,0,0,0}, /* ц */
        {0,0,0,0,0,0,0,3,1,0,0,0,0,0,0,0,0,0,1,1,0,0,1,0,9,3,3,0,0,0,0,0,0}, /* ч */
        {0,0,0,0,0,0,0,3,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,9,3,0,0,0,0,0,0}, /* ш */
        {0,0,0,0,0,0,0,3,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,3,9,0,0,0,0,0,0}, /* щ */
        {0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,9,0,3,0,0,0}, /* ъ */
        {1,0,0,0,0,1,1,0,0,2,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,9,0,1,1,1}, /* ы */
        {0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,0,9,0,0,0}, /* ь */
        {1,0,0,0,0,3,3,0,0,1,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,1,0,9,1,1}, /* э */
        {1,0,0,0,0,1,1,0,0,1,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,0,0,0,1,0,1,9,1}, /* ю */
        {2,0,0,0,0,1,1,0,0,1,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,1,0,1,1,9}  /* я */
      /* а б в г д е ё ж з и й к л м н о п р с т у ф х ц ч ш щ ъ ы ь э ю я */
    };

    String lo_letters="абвгдеёжзийклмнопрстуфхцчшщъыьэюя";

    public FreshEyeProcessor( XText xText, int Threshold_val ) {
        super( xText );
        voc_except.add( "белым бело" );
        voc_except.add( "больше больше" );
        voc_except.add( "больше более" );
        voc_except.add( "более больше" );
        voc_except.add( "больше меньше" );
        voc_except.add( "бы бы" );
        voc_except.add( "бы был" );
        voc_except.add( "бы была" );
        voc_except.add( "бы были" );
        voc_except.add( "бы было" );
        voc_except.add( "бы вы" );
        voc_except.add( "был бы" );
        voc_except.add( "была бы" );
        voc_except.add( "были бы" );
        voc_except.add( "было бы" );
        voc_except.add( "вины виноватый" );
        voc_except.add( "волей неволей" );
        voc_except.add( "время времени" );
        voc_except.add( "всего навсего" );
        voc_except.add( "вы бы" );
        voc_except.add( "даже уже" );
        voc_except.add( "друг друга" );
        voc_except.add( "друг друге" );
        voc_except.add( "друг другом" );
        voc_except.add( "друг другу" );
        voc_except.add( "дурак дураком" );
        voc_except.add( "если если" );
        voc_except.add( "звонка звонка" );
        voc_except.add( "или или" );
        voc_except.add( "как так" );
        voc_except.add( "конце концов" );
        voc_except.add( "корки корки" );
        voc_except.add( "кто что" );
        voc_except.add( "либо либо" );
        voc_except.add( "мало помалу" );
        voc_except.add( "меньше больше" );
        voc_except.add( "начать сначала" );
        voc_except.add( "не на" );
        voc_except.add( "не не" );
        voc_except.add( "не ни" );
        voc_except.add( "него нет" );
        voc_except.add( "ни на" );
        voc_except.add( "ни не" );
        voc_except.add( "ни ни" );
        voc_except.add( "но на" );
        voc_except.add( "но не" );
        voc_except.add( "но ни" );
        voc_except.add( "новые новые" );
        voc_except.add( "объять необъятное" );
        voc_except.add( "одному тому" );
        voc_except.add( "полным полно" );
        voc_except.add( "постольку поскольку" );
        voc_except.add( "так как" );
        voc_except.add( "тем чем" );
        voc_except.add( "то то" );
        voc_except.add( "тогда когда" );
        voc_except.add( "ха ха" );
        voc_except.add( "чем тем" );
        voc_except.add( "что то" );
        voc_except.add( "чуть чуть" );
        voc_except.add( "шаг шагом" );
        voc_except.add( "этой что" );
        voc_except.add( "этот что" );
        this.Threshold_val = Threshold_val;
        for ( int i = 0; i < CONTEXT_LENGTH; i++ ) {
            psychlen[i] = 0;
            wordlist[i] = new WordStruc( null, "", 0, false );
        }
    }

    private int implen( int x ) {
        if ( x == 2 ) return 5;
        int t = ( x - 1 ) / 6;
        return x - t * t + (int)(4.1 / (float)x);
    }
    private int infor( String a, String b ) {
        int count = 0;
        int res = 0;
        Boolean first_elem = true;
        int p;
        int pp = 0;
        int alen = a.length();
        int blen = b.length();
        while ( pp < alen ) {
            if ( (p = b.indexOf( a.charAt( pp ) )) != -1 ) {
                if ( first_elem && ( p == 0 ) ) {
                    res += inf_letters[ lo_letters.indexOf( a.charAt( pp ) ) ][ 1 ];
                } else {
                    res += inf_letters[ lo_letters.indexOf( a.charAt( pp ) ) ][ 0 ];
                }
                count++;
            }
            first_elem = false;
            pp++;
        }
        pp = 0;
        while ( pp < alen ) {
            if ( ( p = b.indexOf( a.charAt( pp ) ) ) == -1 ) {
                if ( pp == 0 ) {
                    res += 2000 - inf_letters[ lo_letters.indexOf( a.charAt( pp ) ) ][ 1 ];
                } else {
                    res += 2000 - inf_letters[ lo_letters.indexOf( a.charAt( pp ) ) ][ 0 ];
                }
                count ++;
            }
            pp++;
        }
        pp = 0;
        while ( pp < blen ) {
            if ( ( p = a.indexOf( b.charAt( pp ) ) ) == -1 ) {
                if ( pp == 0 ) {
                    res += 2000 - inf_letters[ lo_letters.indexOf( b.charAt( pp ) ) ][ 1 ];
                } else {
                    res += 2000 - inf_letters[ lo_letters.indexOf( b.charAt( pp ) ) ][ 0 ];
                }
                count ++;
            }
            pp++;
        }
        if ( count != 0 ) {
            return res / count;
        } else {
            return 0;
        }
    }


    private int SimWords( String a, String b ) {
        int ta = 0, tb = 0;
        int tx = 0, ty = 0;
        int res = 0, resa = 0;
        Boolean rever = false;
        int partlen;
        int prir, dist;

        if ( voc_except.contains( a + " " + b ) ) {
            return 0;
        }
        /*
        if ( a.equals( "подбора" ) && b.equals( "подробно" ) )  {
            String dummy = a;
        }
        */
        if ( a.length() > b.length() ) { // a must be always the shortest
            String tmp;
            rever = true;
            tmp = a;
            a = b;
            b = tmp;
        }
        // optimization vars
        int cparta, cpartb;
        int alen = a.length();
        int blen = b.length();
        int alenmul3 = alen * 3;
        int blenmul3 = blen * 3;
        int basedist = 3 * ( alen + blen ) / 8 + 1;
        int[] a_lnum = new int[ alen ];
        int[] b_lnum = new int[ blen ];
        cparta = alen - 1;
        cpartb = blen - 1;
        for (ta = 0; ta < alen; ta++) {
            a_lnum[ ta ] = lo_letters.indexOf( a.charAt( ta ) );
        }
        for (tb = 0; tb < blen; tb++) {
            b_lnum[ tb ] = lo_letters.indexOf( b.charAt( tb ) );
        }
        // end of optimization vars
        if ( rever ) {
            for ( partlen = 1; partlen <= alen; partlen++, resa = 0 ) {
                for ( ta = 0; ta <= cparta; ta++ ) {
                    int[] parta = new int[ partlen ];
                    System.arraycopy( a_lnum, ta, parta, 0, partlen );
                    for ( tb = 0; cpartb >= tb; tb++ ) {
                        for ( prir = tx = 0, ty = tb; tx < partlen; tx++, ty++ ) prir += sim_ch[ parta[ tx ] ][ b_lnum[ ty ] ];
                        if ( prir != 0 ) {
                            if ( ta > 0 ) prir -= ( prir * ta ) / alenmul3;
                            if ( tb > 0 ) prir -= ( prir * tb ) / blenmul3;
                            dist = blen - ( tb + partlen ) + ta;
                            if ( dist < 3 ) prir += ( prir * ( 2 - dist ) ) / 3;
                            if ( prir > resa ) resa = prir;
                        }
                    }
                }
                if ( ( resa / partlen ) > 6 ) {
                    prir = resa;
                    dist = basedist;
                    res += resa + prir * ( partlen - Math.min( dist, alen ) ) / ( 2 * dist );
                }
                cparta--; cpartb--;
            }
        } else {
            for ( partlen = 1; partlen <= alen; partlen++, resa = 0 ) {
                for ( ta = 0; ta <= cparta; ta++ ) {
                    int[] parta = new int[ partlen ];
                    System.arraycopy( a_lnum, ta, parta, 0, partlen);
                    for ( tb = 0; cpartb >= tb; tb++ ) {
                        for ( prir = tx = 0, ty = tb; tx < partlen; tx++, ty++ ) prir += sim_ch[ parta[ tx ] ][ b_lnum[ ty ] ];
                        if ( prir != 0 ) {
                            if ( ta > 0 ) prir -= ( prir * ta ) / alenmul3;
                            if ( tb > 0 ) prir -= ( prir * tb ) / blenmul3;
                            dist = alen - ( ta + partlen ) + tb;
                            if ( dist < 3 ) prir += ( prir * ( 2 - dist ) ) / 3;
                            if ( prir > resa ) resa = prir;
                        }
                    }
                }
                if ( ( resa / partlen ) > 6 ) {
                    prir = resa;
                    dist = basedist;
                    res += resa + prir * ( partlen - Math.min( dist, alen ) ) / ( 2 * dist );
                }
                cparta--; cpartb--;
            }
        }
        for ( partlen = 1, resa = 0; partlen <= alen; partlen++ ) {
            resa += 9 * partlen;
        }
        res = ( res * infor( a, b ) ) / resa;
        res -= ( res * ( blen - alen ) ) / ( 2 * blen );
/*
        var impa = implen( alen );
        var impb = implen( blen );
 */
        return ( res * alen * blen ) / ( implen( alen ) * implen( blen ) );
    }

    // returns whether the current wrd is similar to one of the previous words from wordlist ("former" w)
    private Object CheckWord( String wrd ) {
        int similarity;
        int t1, dist;
        int badness;
        double dal;
        for ( int t = 0; t < CONTEXT_LENGTH; t++ ) {
            if ( !wordlist[t].word.equals( "" ) ) {
                similarity = SimWords( wordlist[t].word, wrd );
                if ( similarity == 0 ) {
                    continue;
                }
                for ( t1 = t, dist = 0; t1 < CONTEXT_LENGTH; t1++ ) {
                    dist += psychlen[ t1 ];
                }
                t1 = t + 1;
                while ( t1 < CONTEXT_LENGTH ) {
                    dist += wordlist[ t1++ ].word.length() / 3 + 1;
                }
                dal = Math.exp( (double)(-dist * dist) / twosigmasqr );
                badness = (int)( similarity * dal );
                if ( badness > Threshold_val ) {
                    int result[] = {t, badness};
                    // t is the index of worldlist to highlight
                    // badness can be potentially used for statistics
                    return result;
                }
            }
        }
        return false;
    }

    private void ShiftWordQueue( WordStruc wrd, int ra ) {
        for ( int t = 0; t < CONTEXT_LENGTH - 1; t++ ) {
            wordlist[ t ] = wordlist[ t + 1 ];
            psychlen[ t ] = psychlen[ t + 1 ];
        }
        wordlist[ CONTEXT_LENGTH - 1 ] = wrd;
        psychlen[ CONTEXT_LENGTH - 1 ] = ra;
    }

    // originally known as raz() - psychological length of separators
    private int SepPsychoLength( char c ) {
        int res = 0;
        switch( c ) {
        case ',':
            res += 2;
            break;
        case '.': case '!': case '?':
            res += 4;
            break;
        case ';': case ':': case '(': case ')': case '"':
            res +=3;
            break;
        case '-':
            if ( spaces ) {
                res +=3;
            } else {
                res++;
            }
            break;
        default: // ' ' and other chars
            if ( !spaces ) {
                res++;
                spaces = true;
            }
        }
        return res;
    }

    @Override public void ProcessParagraph( XTextRange Para ) {
        super.ProcessParagraph( Para );
        // process paragraph
        String s = Para.getString();
        String cs;
/*
        if ( !s.matches( REGEXP_LETTER ) )
            return;
 */
        StringBuffer word = new StringBuffer( "" ); // current word
        String wordStr;
        char c;
        Object hilite;
        int ra = 0;
        int sp = 0;
        Boolean parbegin = true; // "former" newline
        spaces = false;
        while ( sp < s.length() ) {
            c = s.charAt( sp );
            if ( Character.toString( c ).matches( REGEXP_LETTER ) ) {
                word = new StringBuffer( "" );
                do {
                    word.append( c );
                    if ( ++sp >= s.length() )
                        break;
                    c = s.charAt( sp );
                } while ( Character.toString( c ).matches( REGEXP_LETTER ) );
                // store word data
                wordStr = word.toString().toLowerCase();
                spaces = false;
                psychlen[ CONTEXT_LENGTH - 1 ] = ra;
                hilite = CheckWord( wordStr );
                WordStruc wordstruc = new WordStruc( Para, wordStr, sp - wordStr.length(), false );
                if ( !(hilite instanceof Boolean) ) {
                    wordstruc.hilite();
                    wordlist[ ((int[])hilite)[0] ].hilite();
                }
                ShiftWordQueue( wordstruc, ra );
            } else {
                ra = 0;
                do {
                    ra += SepPsychoLength( c );
                    if ( ++sp >= s.length() )
                        break;
                    c = s.charAt( sp );
                } while ( !Character.toString( c ).matches( REGEXP_LETTER ) );
                if ( !firstpara && parbegin ) {
                    ra += 8;
                    parbegin = false;
                }
            }
        }
        firstpara = false;
    }

}
