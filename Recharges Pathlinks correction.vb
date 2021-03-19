=IF((LEFT(K2,2)="BK"),"GRN",L2)

=IF((LEFT(K2,2)="BK"),"GRN",(IF((LEFT(K2,2)="BL"),"LCN",(IF((LEFT(K2,2)="BB"),"BOS",(IF((LEFT(K2,2)="BG"),"POW",L2))))))))