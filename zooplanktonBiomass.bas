'All lengths are in mm
'Resulting biomass is in mg

Function Bottrell(x, y, L As Double) As Double
    Bottrell = Exp(x + y * Log(L)) / 1000
End Function

Public Function zpBiomassa(length As Double, taxon As String) As Double

zpBiomassa = -1

Select Case taxon:
        Case "bosmina", "bosmina longirostris", "bosmina longispina", "bosmina coregoni":
                zpBiomassa = Bottrell(2.7116, 2.5294, length)
        Case "daphnia", "daphnia sp.", "daphnia longispina", "daphnia cristata", "daphnia galeata", "daphnia cucullata":
                zpBiomassa = Bottrell(1.6, 2.84, length)
        Case "calanoid", "diaptomus gracilis", "calanoid copepodit", "eurytemora sp.", "diaptomus gracilis hane", "limnocalanus marcurus", "eurytemora hane":
                zpBiomassa = Bottrell(2.0365, 1.8911, length)
        Case "cyclopoid", "cyclops scutifer hona", "cyclops scutifer hane", "cyclopoid copepodit":
                zpBiomassa = Bottrell(1.9526, 2.399, length)
        Case "holopedium":
                zpBiomassa = Bottrell(5.3976, 2.0555, length)
        Case "sida":
                zpBiomassa = Bottrell(1.5163, 2.7515, length)
        Case "diaphanosoma", "diaphanosoma brachyorum":
                zpBiomassa = Bottrell(1.6242, 3.0468, length)
        Case "chydorus", "chydorus sphaericus", "eurycercus lamellatus":
                zpBiomassa = Bottrell(2.77, 3.84, length)
        Case "allonella":
                zpBiomassa = Bottrell(2.77, 3.84, length)
        Case "polyphemus":
                zpBiomassa = Bottrell(1.9363, 2.15, length)
        Case "ceriodaphnia", "ceriodaphnia quadrangula":
                zpBiomassa = Bottrell(2.5623, 3.338, length)
        Case "nauplii", "cyclopoid nauplii", "calanoid nauplii":
                zpBiomassa = Bottrell(0.951657875711446, 1.6349, length)
        Case "chaoboridae", "chaoborus flavicans":
                zpBiomassa = (0.1127 * (length ^ 3.2345)) / 1000  'Dumont & Balvay, 1979, Hydrobiologia p. 142 
        Case "leptodora":
                zpBiomassa = Bottrell(-0.821, 2.67, length)
                
        ''''Simplified formula for Rotifera from H. H. Bottrell, p. 452
        
        Case "aneuropsis", "aneuropsis fissa":
                zpBiomassa = 0.03 * length ^ 3 * 0.1
        Case "ascomorpha", "ascomorpha saltans", "ascomorpha ecaudis", "ascomorpha ovalis":
                zpBiomassa = 0.1 * length ^ 3 * 0.1
        Case "asplanchna", "asplanchna sp.", "asplanchna priodonta":
                zpBiomassa = 0.23 * length ^ 3 * 0.039
        Case "brachionus", "brachionus sp.", "brachionus urceolaris":
                zpBiomassa = 0.12 * length ^ 3 * 0.1
        Case "conochilus", "conochilus sp.", "conochilus unicornis", "conochilus hippocrepis":
                zpBiomassa = 0.26 * length * (length * 0.72599901) ^ 2 * 0.1   '=(0,26*(C26/$AG26)*(POWER((C27/$AG27);2)))*0,1
        Case "collotheca":
                zpBiomassa = (length / 7) ^ 3 * 0.1
        Case "euchlanis", "euchlanis dilatata":
                zpBiomassa = 0.1 * length ^ 3 * 0.1
        Case "filina", "filinia sp.", "filinia longiseta":
                zpBiomassa = 0.13 * length ^ 3 * 0.1
        Case "gastropus", "gastropus stylifer":
                zpBiomassa = 0.2 * length ^ 3 * 0.1
        Case "hexartha":
                zpBiomassa = 0.13 * length ^ 3 * 0.1
        Case "kellicottia", "kelikottia longispina":
                zpBiomassa = 0.03 * length ^ 3 * 0.1
        Case "keratella quadrata":
                zpBiomassa = 0.22 * length ^ 3 * 0.1
        Case "keratella cochlearis", "keratella tecta", "keratella hispida":
                zpBiomassa = 0.02 * length ^ 3 * 0.1
        Case "notholca":
                zpBiomassa = 0.035 * length ^ 3 * 0.1
        Case "ploesoma hudsoni":
                zpBiomassa = 0.1 * length ^ 3 * 0.1
        Case "ploesoma triacanthum":
                zpBiomassa = 0.23 * length ^ 3 * 0.1
        Case "polyarthra", "polyarthra sp.", "polyarthra vulgaris", "polyarthra dolicopthera", "polyarthra remata":
                zpBiomassa = 0.28 * length ^ 3 * 0.1
        Case "pompholyx", "pompholyx sulcata":
                zpBiomassa = 0.15 * length ^ 3 * 0.1
        '''Note that Lecane is not specified in Bottrell
        Case "lecane", "lecane sp.", "lecane luna", "lecane lunaris", "lecane crepida": 'Approximated as same as for Pompholyx
                zpBiomassa = 0.15 * length ^ 3 * 0.1
        Case "synchaeta", "synchaeta sp.", "synchaeta grandis":
                zpBiomassa = 0.1 * length ^ 3 * 0.1
        Case "testudinella":
                zpBiomassa = 0.08 * length ^ 3 * 0.1
        Case "trichocerca", "trichocera porcellus", "trichocerca longiseta", "trichocerca capucina", "trichocerca similis", "trichocerca porcellus", "trichocerca pusilla", "trichocerca cylindrica":
                zpBiomassa = 0.52 * length * (0.57 * length) ^ 2 * 0.1 '0.57
                
        Case Else:
            
            
        End Select

End Function
