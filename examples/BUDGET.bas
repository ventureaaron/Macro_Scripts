Attribute VB_Name = "Module1"
Sub budget_categorize()
'this was setup for an export file from my bank. It would first set up an area to count the totals for each category and then calculate the totals
'it should be noted that this is bad programming at it's finest, the better way would have been instead of "hard coding" many different types for
'this macro to check on and then decide the category, a secondary cross walk spreadsheet would have reduced these many lines to just a few lines that
'would apply more generally to anyone that could make a few mods to this script and would then only need a secondary spreadsheet.
'hindsight is even better than 20/20 in my opinion.

    Range("k1").Value = "Expected:"
    Range("l1").Value = "0"
    Range("m1").Value = "Exp. CAR:"
    Range("n1").Value = "0"
    Range("o1").Value = "Exp. CCARD:"
    Range("p1").Value = "0"
    Range("q1").Value = "Exp. ENT:"
    Range("r1").Value = "0"
    Range("s1").Value = "Exp. INS:"
    Range("t1").Value = "0"
    Range("u1").Value = "Exp. MED:"
    Range("v1").Value = "0"
    Range("w1").Value = "Exp. MORT:"
    Range("x1").Value = "0"
    Range("y1").Value = "Exp. SCH:"
    Range("z1").Value = "0"
    Range("aa1").Value = "Exp. UTIL:"
    Range("ab1").Value = "0"
    Range("ac1").Value = "UNEXPECTED:"
    Range("ad1").Value = "0"
    Range("ae1").Value = "Un. ENT:"
    Range("af1").Value = "0"
    Range("ag1").Value = "Un. FFOOD:"
    Range("ah1").Value = "0"
    Range("ai1").Value = "Un. FUEL:"
    Range("aj1").Value = "0"
    Range("ak1").Value = "Un. GROC.:"
    Range("al1").Value = "0"
    Range("am1").Value = "Un. MED:"
    Range("an1").Value = "0"
    Range("ao1").Value = "Un. MISC:"
    Range("ap1").Value = "0"
    Range("aq1").Value = "Un. SUPPLIES:"
    Range("ar1").Value = "0"
    Range("as1").Value = "Un. AUTO:"
    Range("at1").Value = "0"
    Range("au1").Value = "Un. TRAVEL:"
    Range("av1").Value = "0"
    Range("g2").Select
    Do Until Selection.Value = ""
		if Selection.Offset(0,2).Value = "" Then
			If InStr(1, Selection.Value, "NETFLIX") Then
				Selection.Offset(0, 2).Value = "Exp. ENTERTAINMENT"
				Range("r1").Value = (Range("r1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "STARBUCK") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ALBERTSONS") Then
				Selection.Offset(0, 2).Value = "Un. GROCERIES"
				Range("al1").Value = (Range("al1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CITI CARD") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "LORI'S CAFE") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "PIZZA") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "HEB") Then
				Selection.Offset(0, 2).Value = "Un. GROCERIES"
				Range("al1").Value = (Range("al1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "HEB GAS") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CHASE") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "STUDENT LN") Then
				Selection.Offset(0, 2).Value = "Exp. SCHOOL"
				Range("z1").Value = (Range("z1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Starbucks") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ROUNDPOINT") Then
				Selection.Offset(0, 2).Value = "Exp. MORTGAGE"
				Range("x1").Value = (Range("x1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "SAM") Then
				Selection.Offset(0, 2).Value = "Un. GROCERIES"
				Range("al1").Value = (Range("al1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "TARGET") Then
				Selection.Offset(0, 2).Value = "Un. SUPPLIES"
				Range("ar1").Value = (Range("ar1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CAPITAL ONE") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Synchrony") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "VZ WIRELESS") Then
				Selection.Offset(0, 2).Value = "Exp. UTIL"
				Range("ab1").Value = (Range("ab1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "DISCOVER") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ATT PAY") Then
				Selection.Offset(0, 2).Value = "Exp. UTIL"
				Range("ab1").Value = (Range("ab1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Reliant Energy") Then
				Selection.Offset(0, 2).Value = "Exp. UTIL"
				Range("ab1").Value = (Range("ab1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ALLY") Then
				Selection.Offset(0, 2).Value = "Exp. CAR"
				Range("n1").Value = (Range("n1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "UT BOOKSTORE") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "7-ELEVEN") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "TACO") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Audible") Then
				Selection.Offset(0, 2).Value = "Un. ENTERTAINMENT"
				Range("af1").Value = (Range("af1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Amazon Video") Then
				Selection.Offset(0, 2).Value = "Un. ENTERTAINMENT"
				Range("af1").Value = (Range("af1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CASA LOPEZ") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "DILLARDS") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "GRUB BURG") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "WAL-MART") Then
				Selection.Offset(0, 2).Value = "Un. SUPPLIES"
				Range("ar1").Value = (Range("ar1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "MICHAELS STORES") Then
				Selection.Offset(0, 2).Value = "Un. SUPPLIES"
				Range("ar1").Value = (Range("ar1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "WHATABURGER") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ARBYS") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "PILOT") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CAPEXMD") Then
				Selection.Offset(0, 2).Value = "Exp. MED"
				Range("v1").Value = (Range("v1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "JUMBURRITO") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "PANERA") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "COST PLUS") Then
				Selection.Offset(0, 2).Value = "Un. SUPPLIES"
				Range("ar1").Value = (Range("ar1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "FORT WORTH FERT") Then
				Selection.Offset(0, 2).Value = "Un. MED"
				Range("an1").Value = (Range("an1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Hulu") Then
				Selection.Offset(0, 2).Value = "Exp. ENTERTAINMENT"
				Range("r1").Value = (Range("r1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "SONIC") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "LUIGIS") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "TEXAS BURGER") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "STRIPES") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "FIREHOUSE") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "MICHAELS CHAR") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "SHELL") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "PHARMACY") Then
				Selection.Offset(0, 2).Value = "Un. MED"
				Range("an1").Value = (Range("an1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "BEST BUY PAY") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Nelnet") Then
				Selection.Offset(0, 2).Value = "Exp. SCHOOL"
				Range("z1").Value = (Range("z1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ALLSUPS") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CARL'S JR") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ATMOS ENERGY") Then
				Selection.Offset(0, 2).Value = "Exp. UTIL"
				Range("ab1").Value = (Range("ab1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "MINI MART") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "DICKEYS") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Lowes CC") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "PROG COUNTY") Then
				Selection.Offset(0, 2).Value = "Exp. INSURANCE"
				Range("t1").Value = (Range("t1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ROSA'S") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "PAPA JOHN") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "SCHLOTZSKY") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "LOVE") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CHEVRON") Then
				Selection.Offset(0, 2).Value = "Un. FUEL"
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Monitronics") Then
				Selection.Offset(0, 2).Value = "Exp. UTIL"
				Range("ab1").Value = (Range("ab1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "SIRIUSXM") Then
				Selection.Offset(0, 2).Value = "Exp. ENTERTAINMENT"
				Range("r1").Value = (Range("r1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "TEA2GO") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ADVANCE AUTO") Then
				Selection.Offset(0, 2).Value = "Un. AUTO"
				Range("at1").Value = (Range("at1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "COMFORT SUITES") Then
				Selection.Offset(0, 2).Value = "Un. TRAVEL"
				Range("av1").Value = (Range("av1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CHARTWELLS") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "JERSEY GIRL") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "MIDLAND WOMENS") Then
				Selection.Offset(0, 2).Value = "Un. MED"
				Range("an1").Value = (Range("an1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "POPEYE") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "WESTLAKE") Then
				Selection.Offset(0, 2).Value = "Un. SUPPLIES"
				Range("ar1").Value = (Range("ar1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "OREILLY") Then
				Selection.Offset(0, 2).Value = "Un. AUTO"
				Range("at1").Value = (Range("at1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "JACK IN THE") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "LOWES #00063") Then
				Selection.Offset(0, 2).Value = "Un. SUPPLIES"
				Range("ar1").Value = (Range("ar1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "DONUTS") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "DQ") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "COMIC") Then
				Selection.Offset(0, 2).Value = "Un. MISC"
				Range("ap1").Value = (Range("ap1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "HALF PRICE BOOK") Then
				Selection.Offset(0, 2).Value = "Un. MISC"
				Range("ap1").Value = (Range("ap1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "GEICO") Then
				Selection.Offset(0, 2).Value = "Exp. INSURANCE"  
				Range("t1").Value = (Range("t1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ROCKIN Q") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "MCALISTERS") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CHICK-FIL") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "PETSMART") Then
				Selection.Offset(0, 2).Value = "Un. GROCERIES"    
				Range("al1").Value = (Range("al1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "SECURITY FINANC") Then
				Selection.Offset(0, 2).Value = "Exp. CCARD"
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "TEXAS REPRODUCT") Then
				Selection.Offset(0, 2).Value = "Un. MED"   
				Range("an1").Value = (Range("an1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "Amazon Digital") Then
				Selection.Offset(0, 2).Value = "Un. ENTERTAINMENT"     
				Range("af1").Value = (Range("af1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "AMAZON DIG") Then
				Selection.Offset(0, 2).Value = "Un. ENTERTAINMENT"    
				Range("af1").Value = (Range("af1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "AMAZON.COM") Then
				Selection.Offset(0, 2).Value = "Un. MISC"    
				Range("ap1").Value = (Range("ap1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CINERGY") Then
				Selection.Offset(0, 2).Value = "Un. ENTERTAINMENT"
				Range("af1").Value = (Range("af1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "LASERWAS") Then
				Selection.Offset(0, 2).Value = "Un. MISC"    
				Range("ap1").Value = (Range("ap1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "ABUELOS") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "BUDGET RENT") Then
				Selection.Offset(0, 2).Value = "Un. TRAVEL"
				Range("av1").Value = (Range("av1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "BUNDT") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "WHICH WICH") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "FAZOLIS") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "BURGER") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "GOLDEN SUZUKI") Then
				Selection.Offset(0, 2).Value = "Un. ENTERTAINMENT"
				Range("af1").Value = (Range("af1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "WM SUPERCENTER") Then
				Selection.Offset(0, 2).Value = "Un. GROCERIES"    
				Range("al1").Value = (Range("al1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "SPARKFUN") Then
				Selection.Offset(0, 2).Value = "Un. MISC"    
				Range("ap1").Value = (Range("ap1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "DINER") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "CHILI'S") Then
				Selection.Offset(0, 2).Value = "Un. FASTFOOD"    
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			ElseIf InStr(1, Selection.Value, "PETCO") Then
				Selection.Offset(0, 2).Value = "Un. GROCERIES"    
				Range("al1").Value = (Range("al1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
			
			End If
        elseif Selection.Offset(0,2).Value = "Exp. CAR" then
				Range("n1").Value = (Range("n1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Exp. CCARD" then
				Range("p1").Value = (Range("p1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Exp. ENTERTAINMENT" then
				Range("r1").Value = (Range("r1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Exp. INSURANCE" then
				Range("t1").Value = (Range("t1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Exp. MED" then
				Range("v1").Value = (Range("v1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)	
		elseif Selection.Offset(0,2).Value = "Exp. MORTGAGE" then
				Range("x1").Value = (Range("x1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Exp. SCHOOL" then
				Range("z1").Value = (Range("z1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Exp. UTIL" then
				Range("ab1").Value = (Range("ab1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Un. ENTERTAINMENT" then
				Range("af1").Value = (Range("af1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)	
		elseif Selection.Offset(0,2).Value = "Un. FASTFOOD" then
				Range("ah1").Value = (Range("ah1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Un. FUEL" then
				Range("aj1").Value = (Range("aj1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Un. GROCERIES" then
				Range("al1").Value = (Range("al1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)	
		elseif Selection.Offset(0,2).Value = "Un. MED" then
				Range("an1").Value = (Range("an1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)	
		elseif Selection.Offset(0,2).Value = "Un. MISC" then
				Range("ap1").Value = (Range("ap1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)	
		elseif Selection.Offset(0,2).Value = "Un. SUPPLIES" then
				Range("ar1").Value = (Range("ar1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)
		elseif Selection.Offset(0,2).Value = "Un. AUTO" then
				Range("at1").Value = (Range("at1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)	
		elseif Selection.Offset(0,2).Value = "Un. TRAVEL" then
				Range("av1").Value = (Range("av1").Value-Selection.Offset(0,-4).Value+Selection.Offset(0,-3).Value)		
		end if
        Selection.Offset(1, 0).Select
    Loop
    Range("l1").Value= Range("n1").Value + Range("p1").Value +Range("r1").Value +Range("t1").Value +Range("v1").Value +Range("x1").Value +Range("z1").Value +Range("ab1").Value
    Range("ad1").Value= Range("af1").Value + Range("ah1").Value +Range("aj1").Value +Range("al1").Value +Range("an1").Value +Range("ap1").Value +Range("ar1").Value +Range("at1").Value +Range("av1").Value
    Range("a1").Select
    
End Sub
