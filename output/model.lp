Maximize
obj: IV_1 = {1..162}

Subject To
REIF_3 = [0,1]=>ARITHM ([Vá.1->Vá.2::0 >= 1]), !REIF_3 = [0,1]=>ARITHM ([Vá.1->Vá.2::0 <= 0])
REIF_4 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_4 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_4.GEQ.REIF_3)])
REIF_5 = [0,1]=>ARITHM ([Vá.1->Vá.2::1 >= 1]), !REIF_5 = [0,1]=>ARITHM ([Vá.1->Vá.2::1 <= 0])
REIF_6 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_6 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_6.GEQ.REIF_5)])
REIF_7 = [0,1]=>ARITHM ([Vá.1->Vá.2::2 >= 1]), !REIF_7 = [0,1]=>ARITHM ([Vá.1->Vá.2::2 <= 0])
REIF_8 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_8 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_8.GEQ.REIF_7)])
REIF_9 = [0,1]=>ARITHM ([Vá.1->Vá.2::3 >= 1]), !REIF_9 = [0,1]=>ARITHM ([Vá.1->Vá.2::3 <= 0])
REIF_10 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_10 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_10.GEQ.REIF_9)])
REIF_11 = [0,1]=>ARITHM ([Vá.1->Vá.2::4 >= 1]), !REIF_11 = [0,1]=>ARITHM ([Vá.1->Vá.2::4 <= 0])
REIF_12 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_12 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_12.GEQ.REIF_11)])
REIF_13 = [0,1]=>ARITHM ([Vá.1->Vá.2::5 >= 1]), !REIF_13 = [0,1]=>ARITHM ([Vá.1->Vá.2::5 <= 0])
REIF_14 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_14 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_14.GEQ.REIF_13)])
REIF_15 = [0,1]=>ARITHM ([Vá.1->Vá.2::6 >= 1]), !REIF_15 = [0,1]=>ARITHM ([Vá.1->Vá.2::6 <= 0])
REIF_16 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_16 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_16.GEQ.REIF_15)])
REIF_17 = [0,1]=>ARITHM ([Vá.1->Vá.2::7 >= 1]), !REIF_17 = [0,1]=>ARITHM ([Vá.1->Vá.2::7 <= 0])
REIF_18 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_18 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_18.GEQ.REIF_17)])
REIF_19 = [0,1]=>ARITHM ([Vá.1->Vá.2::8 >= 1]), !REIF_19 = [0,1]=>ARITHM ([Vá.1->Vá.2::8 <= 0])
REIF_20 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_20 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_20.GEQ.REIF_19)])
REIF_21 = [0,1]=>ARITHM ([Vá.1->Vá.2::9 >= 1]), !REIF_21 = [0,1]=>ARITHM ([Vá.1->Vá.2::9 <= 0])
REIF_22 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_22 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_22.GEQ.REIF_21)])
REIF_23 = [0,1]=>ARITHM ([Vá.1->Vá.2::10 >= 1]), !REIF_23 = [0,1]=>ARITHM ([Vá.1->Vá.2::10 <= 0])
REIF_24 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_24 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_24.GEQ.REIF_23)])
REIF_25 = [0,1]=>ARITHM ([Vá.1->Vá.2::11 >= 1]), !REIF_25 = [0,1]=>ARITHM ([Vá.1->Vá.2::11 <= 0])
REIF_26 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_26 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_26.GEQ.REIF_25)])
REIF_27 = [0,1]=>ARITHM ([Vá.1->Vá.2::12 >= 1]), !REIF_27 = [0,1]=>ARITHM ([Vá.1->Vá.2::12 <= 0])
REIF_28 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_28 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_28.GEQ.REIF_27)])
REIF_29 = [0,1]=>ARITHM ([Vá.1->Vá.2::13 >= 1]), !REIF_29 = [0,1]=>ARITHM ([Vá.1->Vá.2::13 <= 0])
REIF_30 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_30 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_30.GEQ.REIF_29)])
REIF_31 = [0,1]=>ARITHM ([Vá.1->Vá.2::14 >= 1]), !REIF_31 = [0,1]=>ARITHM ([Vá.1->Vá.2::14 <= 0])
REIF_32 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_32 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_32.GEQ.REIF_31)])
REIF_33 = [0,1]=>ARITHM ([Vá.1->Vá.2::15 >= 1]), !REIF_33 = [0,1]=>ARITHM ([Vá.1->Vá.2::15 <= 0])
REIF_34 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_34 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_34.GEQ.REIF_33)])
REIF_35 = [0,1]=>ARITHM ([Vá.1->Vá.2::16 >= 1]), !REIF_35 = [0,1]=>ARITHM ([Vá.1->Vá.2::16 <= 0])
REIF_36 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_36 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_36.GEQ.REIF_35)])
REIF_37 = [0,1]=>ARITHM ([Vá.1->Vá.2::17 >= 1]), !REIF_37 = [0,1]=>ARITHM ([Vá.1->Vá.2::17 <= 0])
REIF_38 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_38 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_38.GEQ.REIF_37)])
REIF_39 = [0,1]=>ARITHM ([Vá.1->Vá.2::18 >= 1]), !REIF_39 = [0,1]=>ARITHM ([Vá.1->Vá.2::18 <= 0])
REIF_40 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_40 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_40.GEQ.REIF_39)])
REIF_41 = [0,1]=>ARITHM ([Vá.1->Vá.2::19 >= 1]), !REIF_41 = [0,1]=>ARITHM ([Vá.1->Vá.2::19 <= 0])
REIF_42 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_42 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_42.GEQ.REIF_41)])
REIF_43 = [0,1]=>ARITHM ([Vá.1->Vá.2::20 >= 1]), !REIF_43 = [0,1]=>ARITHM ([Vá.1->Vá.2::20 <= 0])
REIF_44 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_44 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_44.GEQ.REIF_43)])
REIF_45 = [0,1]=>ARITHM ([Vá.1->Vá.2::21 >= 1]), !REIF_45 = [0,1]=>ARITHM ([Vá.1->Vá.2::21 <= 0])
REIF_46 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_46 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_46.GEQ.REIF_45)])
REIF_47 = [0,1]=>ARITHM ([Vá.1->Vá.2::22 >= 1]), !REIF_47 = [0,1]=>ARITHM ([Vá.1->Vá.2::22 <= 0])
REIF_48 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_48 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_48.GEQ.REIF_47)])
REIF_49 = [0,1]=>ARITHM ([Vá.1->Vá.2::23 >= 1]), !REIF_49 = [0,1]=>ARITHM ([Vá.1->Vá.2::23 <= 0])
REIF_50 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_50 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_50.GEQ.REIF_49)])
REIF_51 = [0,1]=>ARITHM ([Vá.1->Vá.2::24 >= 1]), !REIF_51 = [0,1]=>ARITHM ([Vá.1->Vá.2::24 <= 0])
REIF_52 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_52 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_52.GEQ.REIF_51)])
REIF_53 = [0,1]=>ARITHM ([Vá.1->Vá.2::25 >= 1]), !REIF_53 = [0,1]=>ARITHM ([Vá.1->Vá.2::25 <= 0])
REIF_54 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_54 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_54.GEQ.REIF_53)])
REIF_55 = [0,1]=>ARITHM ([Vá.1->Vá.2::26 >= 1]), !REIF_55 = [0,1]=>ARITHM ([Vá.1->Vá.2::26 <= 0])
REIF_56 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_56 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_56.GEQ.REIF_55)])
REIF_57 = [0,1]=>ARITHM ([Vá.1->Vá.2::27 >= 1]), !REIF_57 = [0,1]=>ARITHM ([Vá.1->Vá.2::27 <= 0])
REIF_58 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_58 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_58.GEQ.REIF_57)])
REIF_59 = [0,1]=>ARITHM ([Vá.1->Vá.2::28 >= 1]), !REIF_59 = [0,1]=>ARITHM ([Vá.1->Vá.2::28 <= 0])
REIF_60 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_60 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_60.GEQ.REIF_59)])
REIF_61 = [0,1]=>ARITHM ([Vá.1->Vá.2::29 >= 1]), !REIF_61 = [0,1]=>ARITHM ([Vá.1->Vá.2::29 <= 0])
REIF_62 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_62 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_62.GEQ.REIF_61)])
REIF_63 = [0,1]=>ARITHM ([Vá.1->Vá.2::30 >= 1]), !REIF_63 = [0,1]=>ARITHM ([Vá.1->Vá.2::30 <= 0])
REIF_64 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_64 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_64.GEQ.REIF_63)])
REIF_65 = [0,1]=>ARITHM ([Vá.1->Vá.2::31 >= 1]), !REIF_65 = [0,1]=>ARITHM ([Vá.1->Vá.2::31 <= 0])
REIF_66 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_66 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_66.GEQ.REIF_65)])
REIF_67 = [0,1]=>ARITHM ([Vá.1->Vá.2::32 >= 1]), !REIF_67 = [0,1]=>ARITHM ([Vá.1->Vá.2::32 <= 0])
REIF_68 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_68 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_68.GEQ.REIF_67)])
REIF_69 = [0,1]=>ARITHM ([Vá.1->Vá.2::33 >= 1]), !REIF_69 = [0,1]=>ARITHM ([Vá.1->Vá.2::33 <= 0])
REIF_70 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_70 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_70.GEQ.REIF_69)])
REIF_71 = [0,1]=>ARITHM ([Vá.1->Vá.2::34 >= 1]), !REIF_71 = [0,1]=>ARITHM ([Vá.1->Vá.2::34 <= 0])
REIF_72 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_72 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_72.GEQ.REIF_71)])
REIF_73 = [0,1]=>ARITHM ([Vá.1->Vá.2::35 >= 1]), !REIF_73 = [0,1]=>ARITHM ([Vá.1->Vá.2::35 <= 0])
REIF_74 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_74 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_74.GEQ.REIF_73)])
REIF_75 = [0,1]=>ARITHM ([Vá.1->Vá.2::36 >= 1]), !REIF_75 = [0,1]=>ARITHM ([Vá.1->Vá.2::36 <= 0])
REIF_76 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_76 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_76.GEQ.REIF_75)])
REIF_77 = [0,1]=>ARITHM ([Vá.1->Vá.2::37 >= 1]), !REIF_77 = [0,1]=>ARITHM ([Vá.1->Vá.2::37 <= 0])
REIF_78 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_78 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_78.GEQ.REIF_77)])
REIF_79 = [0,1]=>ARITHM ([Vá.1->Vá.2::38 >= 1]), !REIF_79 = [0,1]=>ARITHM ([Vá.1->Vá.2::38 <= 0])
REIF_80 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_80 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_80.GEQ.REIF_79)])
REIF_81 = [0,1]=>ARITHM ([Vá.1->Vá.2::39 >= 1]), !REIF_81 = [0,1]=>ARITHM ([Vá.1->Vá.2::39 <= 0])
REIF_82 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_82 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_82.GEQ.REIF_81)])
REIF_83 = [0,1]=>ARITHM ([Vá.1->Vá.2::40 >= 1]), !REIF_83 = [0,1]=>ARITHM ([Vá.1->Vá.2::40 <= 0])
REIF_84 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_84 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_84.GEQ.REIF_83)])
REIF_85 = [0,1]=>ARITHM ([Vá.1->Vá.2::41 >= 1]), !REIF_85 = [0,1]=>ARITHM ([Vá.1->Vá.2::41 <= 0])
REIF_86 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_86 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_86.GEQ.REIF_85)])
REIF_87 = [0,1]=>ARITHM ([Vá.1->Vá.2::42 >= 1]), !REIF_87 = [0,1]=>ARITHM ([Vá.1->Vá.2::42 <= 0])
REIF_88 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_88 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_88.GEQ.REIF_87)])
REIF_89 = [0,1]=>ARITHM ([Vá.1->Vá.2::43 >= 1]), !REIF_89 = [0,1]=>ARITHM ([Vá.1->Vá.2::43 <= 0])
REIF_90 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_90 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_90.GEQ.REIF_89)])
REIF_91 = [0,1]=>ARITHM ([Vá.1->Vá.2::44 >= 1]), !REIF_91 = [0,1]=>ARITHM ([Vá.1->Vá.2::44 <= 0])
REIF_92 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_92 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_92.GEQ.REIF_91)])
REIF_93 = [0,1]=>ARITHM ([Vá.1->Vá.2::45 >= 1]), !REIF_93 = [0,1]=>ARITHM ([Vá.1->Vá.2::45 <= 0])
REIF_94 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_94 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_94.GEQ.REIF_93)])
REIF_95 = [0,1]=>ARITHM ([Vá.1->Vá.2::46 >= 1]), !REIF_95 = [0,1]=>ARITHM ([Vá.1->Vá.2::46 <= 0])
REIF_96 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_96 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_96.GEQ.REIF_95)])
REIF_97 = [0,1]=>ARITHM ([Vá.1->Vá.2::47 >= 1]), !REIF_97 = [0,1]=>ARITHM ([Vá.1->Vá.2::47 <= 0])
REIF_98 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_98 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_98.GEQ.REIF_97)])
REIF_99 = [0,1]=>ARITHM ([Vá.1->Vá.2::48 >= 1]), !REIF_99 = [0,1]=>ARITHM ([Vá.1->Vá.2::48 <= 0])
REIF_100 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_100 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_100.GEQ.REIF_99)])
REIF_101 = [0,1]=>ARITHM ([Vá.1->Vá.2::49 >= 1]), !REIF_101 = [0,1]=>ARITHM ([Vá.1->Vá.2::49 <= 0])
REIF_102 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_102 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_102.GEQ.REIF_101)])
REIF_103 = [0,1]=>ARITHM ([Vá.1->Vá.2::50 >= 1]), !REIF_103 = [0,1]=>ARITHM ([Vá.1->Vá.2::50 <= 0])
REIF_104 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_104 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_104.GEQ.REIF_103)])
REIF_105 = [0,1]=>ARITHM ([Vá.1->Vá.2::51 >= 1]), !REIF_105 = [0,1]=>ARITHM ([Vá.1->Vá.2::51 <= 0])
REIF_106 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_106 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_106.GEQ.REIF_105)])
REIF_107 = [0,1]=>ARITHM ([Vá.1->Vá.2::52 >= 1]), !REIF_107 = [0,1]=>ARITHM ([Vá.1->Vá.2::52 <= 0])
REIF_108 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_108 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_108.GEQ.REIF_107)])
REIF_109 = [0,1]=>ARITHM ([Vá.1->Vá.2::53 >= 1]), !REIF_109 = [0,1]=>ARITHM ([Vá.1->Vá.2::53 <= 0])
REIF_110 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_110 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_110.GEQ.REIF_109)])
REIF_111 = [0,1]=>ARITHM ([Vá.1->Vá.2::54 >= 1]), !REIF_111 = [0,1]=>ARITHM ([Vá.1->Vá.2::54 <= 0])
REIF_112 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_112 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_112.GEQ.REIF_111)])
REIF_113 = [0,1]=>ARITHM ([Vá.1->Vá.2::55 >= 1]), !REIF_113 = [0,1]=>ARITHM ([Vá.1->Vá.2::55 <= 0])
REIF_114 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_114 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_114.GEQ.REIF_113)])
REIF_115 = [0,1]=>ARITHM ([Vá.1->Vá.2::56 >= 1]), !REIF_115 = [0,1]=>ARITHM ([Vá.1->Vá.2::56 <= 0])
REIF_116 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_116 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_116.GEQ.REIF_115)])
REIF_117 = [0,1]=>ARITHM ([Vá.1->Vá.2::57 >= 1]), !REIF_117 = [0,1]=>ARITHM ([Vá.1->Vá.2::57 <= 0])
REIF_118 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_118 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_118.GEQ.REIF_117)])
REIF_119 = [0,1]=>ARITHM ([Vá.1->Vá.2::58 >= 1]), !REIF_119 = [0,1]=>ARITHM ([Vá.1->Vá.2::58 <= 0])
REIF_120 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_120 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_120.GEQ.REIF_119)])
REIF_121 = [0,1]=>ARITHM ([Vá.1->Vá.2::59 >= 1]), !REIF_121 = [0,1]=>ARITHM ([Vá.1->Vá.2::59 <= 0])
REIF_122 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_122 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_122.GEQ.REIF_121)])
REIF_123 = [0,1]=>ARITHM ([Vá.1->Vá.2::60 >= 1]), !REIF_123 = [0,1]=>ARITHM ([Vá.1->Vá.2::60 <= 0])
REIF_124 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_124 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_124.GEQ.REIF_123)])
REIF_125 = [0,1]=>ARITHM ([Vá.1->Vá.2::61 >= 1]), !REIF_125 = [0,1]=>ARITHM ([Vá.1->Vá.2::61 <= 0])
REIF_126 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_126 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_126.GEQ.REIF_125)])
REIF_127 = [0,1]=>ARITHM ([Vá.1->Vá.2::62 >= 1]), !REIF_127 = [0,1]=>ARITHM ([Vá.1->Vá.2::62 <= 0])
REIF_128 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_128 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_128.GEQ.REIF_127)])
REIF_129 = [0,1]=>ARITHM ([Vá.1->Vá.2::63 >= 1]), !REIF_129 = [0,1]=>ARITHM ([Vá.1->Vá.2::63 <= 0])
REIF_130 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_130 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_130.GEQ.REIF_129)])
REIF_131 = [0,1]=>ARITHM ([Vá.1->Vá.2::64 >= 1]), !REIF_131 = [0,1]=>ARITHM ([Vá.1->Vá.2::64 <= 0])
REIF_132 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_132 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_132.GEQ.REIF_131)])
REIF_133 = [0,1]=>ARITHM ([Vá.1->Vá.2::65 >= 1]), !REIF_133 = [0,1]=>ARITHM ([Vá.1->Vá.2::65 <= 0])
REIF_134 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_134 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_134.GEQ.REIF_133)])
REIF_135 = [0,1]=>ARITHM ([Vá.1->Vá.2::66 >= 1]), !REIF_135 = [0,1]=>ARITHM ([Vá.1->Vá.2::66 <= 0])
REIF_136 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_136 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_136.GEQ.REIF_135)])
REIF_137 = [0,1]=>ARITHM ([Vá.1->Vá.2::67 >= 1]), !REIF_137 = [0,1]=>ARITHM ([Vá.1->Vá.2::67 <= 0])
REIF_138 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_138 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_138.GEQ.REIF_137)])
REIF_139 = [0,1]=>ARITHM ([Vá.1->Vá.2::68 >= 1]), !REIF_139 = [0,1]=>ARITHM ([Vá.1->Vá.2::68 <= 0])
REIF_140 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_140 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_140.GEQ.REIF_139)])
REIF_141 = [0,1]=>ARITHM ([Vá.1->Vá.2::69 >= 1]), !REIF_141 = [0,1]=>ARITHM ([Vá.1->Vá.2::69 <= 0])
REIF_142 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_142 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_142.GEQ.REIF_141)])
REIF_143 = [0,1]=>ARITHM ([Vá.1->Vá.2::70 >= 1]), !REIF_143 = [0,1]=>ARITHM ([Vá.1->Vá.2::70 <= 0])
REIF_144 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_144 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_144.GEQ.REIF_143)])
REIF_145 = [0,1]=>ARITHM ([Vá.1->Vá.2::71 >= 1]), !REIF_145 = [0,1]=>ARITHM ([Vá.1->Vá.2::71 <= 0])
REIF_146 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_146 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_146.GEQ.REIF_145)])
REIF_147 = [0,1]=>ARITHM ([Vá.1->Vá.2::72 >= 1]), !REIF_147 = [0,1]=>ARITHM ([Vá.1->Vá.2::72 <= 0])
REIF_148 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_148 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_148.GEQ.REIF_147)])
REIF_149 = [0,1]=>ARITHM ([Vá.1->Vá.2::73 >= 1]), !REIF_149 = [0,1]=>ARITHM ([Vá.1->Vá.2::73 <= 0])
REIF_150 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_150 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_150.GEQ.REIF_149)])
REIF_151 = [0,1]=>ARITHM ([Vá.1->Vá.2::74 >= 1]), !REIF_151 = [0,1]=>ARITHM ([Vá.1->Vá.2::74 <= 0])
REIF_152 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_152 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_152.GEQ.REIF_151)])
REIF_153 = [0,1]=>ARITHM ([Vá.1->Vá.2::75 >= 1]), !REIF_153 = [0,1]=>ARITHM ([Vá.1->Vá.2::75 <= 0])
REIF_154 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_154 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_154.GEQ.REIF_153)])
REIF_155 = [0,1]=>ARITHM ([Vá.1->Vá.2::76 >= 1]), !REIF_155 = [0,1]=>ARITHM ([Vá.1->Vá.2::76 <= 0])
REIF_156 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_156 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_156.GEQ.REIF_155)])
REIF_157 = [0,1]=>ARITHM ([Vá.1->Vá.2::77 >= 1]), !REIF_157 = [0,1]=>ARITHM ([Vá.1->Vá.2::77 <= 0])
REIF_158 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_158 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_158.GEQ.REIF_157)])
REIF_159 = [0,1]=>ARITHM ([Vá.1->Vá.2::78 >= 1]), !REIF_159 = [0,1]=>ARITHM ([Vá.1->Vá.2::78 <= 0])
REIF_160 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_160 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_160.GEQ.REIF_159)])
REIF_161 = [0,1]=>ARITHM ([Vá.1->Vá.2::79 >= 1]), !REIF_161 = [0,1]=>ARITHM ([Vá.1->Vá.2::79 <= 0])
REIF_162 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_162 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_162.GEQ.REIF_161)])
REIF_163 = [0,1]=>ARITHM ([Vá.1->Vá.2::80 >= 1]), !REIF_163 = [0,1]=>ARITHM ([Vá.1->Vá.2::80 <= 0])
REIF_164 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_164 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_164.GEQ.REIF_163)])
REIF_165 = [0,1]=>ARITHM ([Vá.2->Vá.1::0 >= 1]), !REIF_165 = [0,1]=>ARITHM ([Vá.2->Vá.1::0 <= 0])
REIF_166 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_166 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_166.GEQ.REIF_165)])
REIF_167 = [0,1]=>ARITHM ([Vá.2->Vá.1::1 >= 1]), !REIF_167 = [0,1]=>ARITHM ([Vá.2->Vá.1::1 <= 0])
REIF_168 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_168 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_168.GEQ.REIF_167)])
REIF_169 = [0,1]=>ARITHM ([Vá.2->Vá.1::2 >= 1]), !REIF_169 = [0,1]=>ARITHM ([Vá.2->Vá.1::2 <= 0])
REIF_170 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_170 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_170.GEQ.REIF_169)])
REIF_171 = [0,1]=>ARITHM ([Vá.2->Vá.1::3 >= 1]), !REIF_171 = [0,1]=>ARITHM ([Vá.2->Vá.1::3 <= 0])
REIF_172 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_172 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_172.GEQ.REIF_171)])
REIF_173 = [0,1]=>ARITHM ([Vá.2->Vá.1::4 >= 1]), !REIF_173 = [0,1]=>ARITHM ([Vá.2->Vá.1::4 <= 0])
REIF_174 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_174 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_174.GEQ.REIF_173)])
REIF_175 = [0,1]=>ARITHM ([Vá.2->Vá.1::5 >= 1]), !REIF_175 = [0,1]=>ARITHM ([Vá.2->Vá.1::5 <= 0])
REIF_176 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_176 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_176.GEQ.REIF_175)])
REIF_177 = [0,1]=>ARITHM ([Vá.2->Vá.1::6 >= 1]), !REIF_177 = [0,1]=>ARITHM ([Vá.2->Vá.1::6 <= 0])
REIF_178 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_178 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_178.GEQ.REIF_177)])
REIF_179 = [0,1]=>ARITHM ([Vá.2->Vá.1::7 >= 1]), !REIF_179 = [0,1]=>ARITHM ([Vá.2->Vá.1::7 <= 0])
REIF_180 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_180 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_180.GEQ.REIF_179)])
REIF_181 = [0,1]=>ARITHM ([Vá.2->Vá.1::8 >= 1]), !REIF_181 = [0,1]=>ARITHM ([Vá.2->Vá.1::8 <= 0])
REIF_182 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_182 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_182.GEQ.REIF_181)])
REIF_183 = [0,1]=>ARITHM ([Vá.2->Vá.1::9 >= 1]), !REIF_183 = [0,1]=>ARITHM ([Vá.2->Vá.1::9 <= 0])
REIF_184 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_184 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_184.GEQ.REIF_183)])
REIF_185 = [0,1]=>ARITHM ([Vá.2->Vá.1::10 >= 1]), !REIF_185 = [0,1]=>ARITHM ([Vá.2->Vá.1::10 <= 0])
REIF_186 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_186 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_186.GEQ.REIF_185)])
REIF_187 = [0,1]=>ARITHM ([Vá.2->Vá.1::11 >= 1]), !REIF_187 = [0,1]=>ARITHM ([Vá.2->Vá.1::11 <= 0])
REIF_188 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_188 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_188.GEQ.REIF_187)])
REIF_189 = [0,1]=>ARITHM ([Vá.2->Vá.1::12 >= 1]), !REIF_189 = [0,1]=>ARITHM ([Vá.2->Vá.1::12 <= 0])
REIF_190 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_190 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_190.GEQ.REIF_189)])
REIF_191 = [0,1]=>ARITHM ([Vá.2->Vá.1::13 >= 1]), !REIF_191 = [0,1]=>ARITHM ([Vá.2->Vá.1::13 <= 0])
REIF_192 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_192 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_192.GEQ.REIF_191)])
REIF_193 = [0,1]=>ARITHM ([Vá.2->Vá.1::14 >= 1]), !REIF_193 = [0,1]=>ARITHM ([Vá.2->Vá.1::14 <= 0])
REIF_194 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_194 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_194.GEQ.REIF_193)])
REIF_195 = [0,1]=>ARITHM ([Vá.2->Vá.1::15 >= 1]), !REIF_195 = [0,1]=>ARITHM ([Vá.2->Vá.1::15 <= 0])
REIF_196 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_196 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_196.GEQ.REIF_195)])
REIF_197 = [0,1]=>ARITHM ([Vá.2->Vá.1::16 >= 1]), !REIF_197 = [0,1]=>ARITHM ([Vá.2->Vá.1::16 <= 0])
REIF_198 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_198 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_198.GEQ.REIF_197)])
REIF_199 = [0,1]=>ARITHM ([Vá.2->Vá.1::17 >= 1]), !REIF_199 = [0,1]=>ARITHM ([Vá.2->Vá.1::17 <= 0])
REIF_200 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_200 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_200.GEQ.REIF_199)])
REIF_201 = [0,1]=>ARITHM ([Vá.2->Vá.1::18 >= 1]), !REIF_201 = [0,1]=>ARITHM ([Vá.2->Vá.1::18 <= 0])
REIF_202 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_202 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_202.GEQ.REIF_201)])
REIF_203 = [0,1]=>ARITHM ([Vá.2->Vá.1::19 >= 1]), !REIF_203 = [0,1]=>ARITHM ([Vá.2->Vá.1::19 <= 0])
REIF_204 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_204 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_204.GEQ.REIF_203)])
REIF_205 = [0,1]=>ARITHM ([Vá.2->Vá.1::20 >= 1]), !REIF_205 = [0,1]=>ARITHM ([Vá.2->Vá.1::20 <= 0])
REIF_206 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_206 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_206.GEQ.REIF_205)])
REIF_207 = [0,1]=>ARITHM ([Vá.2->Vá.1::21 >= 1]), !REIF_207 = [0,1]=>ARITHM ([Vá.2->Vá.1::21 <= 0])
REIF_208 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_208 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_208.GEQ.REIF_207)])
REIF_209 = [0,1]=>ARITHM ([Vá.2->Vá.1::22 >= 1]), !REIF_209 = [0,1]=>ARITHM ([Vá.2->Vá.1::22 <= 0])
REIF_210 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_210 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_210.GEQ.REIF_209)])
REIF_211 = [0,1]=>ARITHM ([Vá.2->Vá.1::23 >= 1]), !REIF_211 = [0,1]=>ARITHM ([Vá.2->Vá.1::23 <= 0])
REIF_212 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_212 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_212.GEQ.REIF_211)])
REIF_213 = [0,1]=>ARITHM ([Vá.2->Vá.1::24 >= 1]), !REIF_213 = [0,1]=>ARITHM ([Vá.2->Vá.1::24 <= 0])
REIF_214 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_214 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_214.GEQ.REIF_213)])
REIF_215 = [0,1]=>ARITHM ([Vá.2->Vá.1::25 >= 1]), !REIF_215 = [0,1]=>ARITHM ([Vá.2->Vá.1::25 <= 0])
REIF_216 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_216 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_216.GEQ.REIF_215)])
REIF_217 = [0,1]=>ARITHM ([Vá.2->Vá.1::26 >= 1]), !REIF_217 = [0,1]=>ARITHM ([Vá.2->Vá.1::26 <= 0])
REIF_218 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_218 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_218.GEQ.REIF_217)])
REIF_219 = [0,1]=>ARITHM ([Vá.2->Vá.1::27 >= 1]), !REIF_219 = [0,1]=>ARITHM ([Vá.2->Vá.1::27 <= 0])
REIF_220 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_220 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_220.GEQ.REIF_219)])
REIF_221 = [0,1]=>ARITHM ([Vá.2->Vá.1::28 >= 1]), !REIF_221 = [0,1]=>ARITHM ([Vá.2->Vá.1::28 <= 0])
REIF_222 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_222 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_222.GEQ.REIF_221)])
REIF_223 = [0,1]=>ARITHM ([Vá.2->Vá.1::29 >= 1]), !REIF_223 = [0,1]=>ARITHM ([Vá.2->Vá.1::29 <= 0])
REIF_224 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_224 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_224.GEQ.REIF_223)])
REIF_225 = [0,1]=>ARITHM ([Vá.2->Vá.1::30 >= 1]), !REIF_225 = [0,1]=>ARITHM ([Vá.2->Vá.1::30 <= 0])
REIF_226 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_226 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_226.GEQ.REIF_225)])
REIF_227 = [0,1]=>ARITHM ([Vá.2->Vá.1::31 >= 1]), !REIF_227 = [0,1]=>ARITHM ([Vá.2->Vá.1::31 <= 0])
REIF_228 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_228 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_228.GEQ.REIF_227)])
REIF_229 = [0,1]=>ARITHM ([Vá.2->Vá.1::32 >= 1]), !REIF_229 = [0,1]=>ARITHM ([Vá.2->Vá.1::32 <= 0])
REIF_230 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_230 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_230.GEQ.REIF_229)])
REIF_231 = [0,1]=>ARITHM ([Vá.2->Vá.1::33 >= 1]), !REIF_231 = [0,1]=>ARITHM ([Vá.2->Vá.1::33 <= 0])
REIF_232 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_232 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_232.GEQ.REIF_231)])
REIF_233 = [0,1]=>ARITHM ([Vá.2->Vá.1::34 >= 1]), !REIF_233 = [0,1]=>ARITHM ([Vá.2->Vá.1::34 <= 0])
REIF_234 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_234 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_234.GEQ.REIF_233)])
REIF_235 = [0,1]=>ARITHM ([Vá.2->Vá.1::35 >= 1]), !REIF_235 = [0,1]=>ARITHM ([Vá.2->Vá.1::35 <= 0])
REIF_236 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_236 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_236.GEQ.REIF_235)])
REIF_237 = [0,1]=>ARITHM ([Vá.2->Vá.1::36 >= 1]), !REIF_237 = [0,1]=>ARITHM ([Vá.2->Vá.1::36 <= 0])
REIF_238 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_238 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_238.GEQ.REIF_237)])
REIF_239 = [0,1]=>ARITHM ([Vá.2->Vá.1::37 >= 1]), !REIF_239 = [0,1]=>ARITHM ([Vá.2->Vá.1::37 <= 0])
REIF_240 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_240 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_240.GEQ.REIF_239)])
REIF_241 = [0,1]=>ARITHM ([Vá.2->Vá.1::38 >= 1]), !REIF_241 = [0,1]=>ARITHM ([Vá.2->Vá.1::38 <= 0])
REIF_242 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_242 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_242.GEQ.REIF_241)])
REIF_243 = [0,1]=>ARITHM ([Vá.2->Vá.1::39 >= 1]), !REIF_243 = [0,1]=>ARITHM ([Vá.2->Vá.1::39 <= 0])
REIF_244 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_244 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_244.GEQ.REIF_243)])
REIF_245 = [0,1]=>ARITHM ([Vá.2->Vá.1::40 >= 1]), !REIF_245 = [0,1]=>ARITHM ([Vá.2->Vá.1::40 <= 0])
REIF_246 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_246 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_246.GEQ.REIF_245)])
REIF_247 = [0,1]=>ARITHM ([Vá.2->Vá.1::41 >= 1]), !REIF_247 = [0,1]=>ARITHM ([Vá.2->Vá.1::41 <= 0])
REIF_248 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_248 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_248.GEQ.REIF_247)])
REIF_249 = [0,1]=>ARITHM ([Vá.2->Vá.1::42 >= 1]), !REIF_249 = [0,1]=>ARITHM ([Vá.2->Vá.1::42 <= 0])
REIF_250 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_250 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_250.GEQ.REIF_249)])
REIF_251 = [0,1]=>ARITHM ([Vá.2->Vá.1::43 >= 1]), !REIF_251 = [0,1]=>ARITHM ([Vá.2->Vá.1::43 <= 0])
REIF_252 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_252 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_252.GEQ.REIF_251)])
REIF_253 = [0,1]=>ARITHM ([Vá.2->Vá.1::44 >= 1]), !REIF_253 = [0,1]=>ARITHM ([Vá.2->Vá.1::44 <= 0])
REIF_254 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_254 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_254.GEQ.REIF_253)])
REIF_255 = [0,1]=>ARITHM ([Vá.2->Vá.1::45 >= 1]), !REIF_255 = [0,1]=>ARITHM ([Vá.2->Vá.1::45 <= 0])
REIF_256 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_256 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_256.GEQ.REIF_255)])
REIF_257 = [0,1]=>ARITHM ([Vá.2->Vá.1::46 >= 1]), !REIF_257 = [0,1]=>ARITHM ([Vá.2->Vá.1::46 <= 0])
REIF_258 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_258 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_258.GEQ.REIF_257)])
REIF_259 = [0,1]=>ARITHM ([Vá.2->Vá.1::47 >= 1]), !REIF_259 = [0,1]=>ARITHM ([Vá.2->Vá.1::47 <= 0])
REIF_260 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_260 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_260.GEQ.REIF_259)])
REIF_261 = [0,1]=>ARITHM ([Vá.2->Vá.1::48 >= 1]), !REIF_261 = [0,1]=>ARITHM ([Vá.2->Vá.1::48 <= 0])
REIF_262 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_262 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_262.GEQ.REIF_261)])
REIF_263 = [0,1]=>ARITHM ([Vá.2->Vá.1::49 >= 1]), !REIF_263 = [0,1]=>ARITHM ([Vá.2->Vá.1::49 <= 0])
REIF_264 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_264 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_264.GEQ.REIF_263)])
REIF_265 = [0,1]=>ARITHM ([Vá.2->Vá.1::50 >= 1]), !REIF_265 = [0,1]=>ARITHM ([Vá.2->Vá.1::50 <= 0])
REIF_266 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_266 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_266.GEQ.REIF_265)])
REIF_267 = [0,1]=>ARITHM ([Vá.2->Vá.1::51 >= 1]), !REIF_267 = [0,1]=>ARITHM ([Vá.2->Vá.1::51 <= 0])
REIF_268 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_268 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_268.GEQ.REIF_267)])
REIF_269 = [0,1]=>ARITHM ([Vá.2->Vá.1::52 >= 1]), !REIF_269 = [0,1]=>ARITHM ([Vá.2->Vá.1::52 <= 0])
REIF_270 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_270 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_270.GEQ.REIF_269)])
REIF_271 = [0,1]=>ARITHM ([Vá.2->Vá.1::53 >= 1]), !REIF_271 = [0,1]=>ARITHM ([Vá.2->Vá.1::53 <= 0])
REIF_272 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_272 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_272.GEQ.REIF_271)])
REIF_273 = [0,1]=>ARITHM ([Vá.2->Vá.1::54 >= 1]), !REIF_273 = [0,1]=>ARITHM ([Vá.2->Vá.1::54 <= 0])
REIF_274 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_274 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_274.GEQ.REIF_273)])
REIF_275 = [0,1]=>ARITHM ([Vá.2->Vá.1::55 >= 1]), !REIF_275 = [0,1]=>ARITHM ([Vá.2->Vá.1::55 <= 0])
REIF_276 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_276 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_276.GEQ.REIF_275)])
REIF_277 = [0,1]=>ARITHM ([Vá.2->Vá.1::56 >= 1]), !REIF_277 = [0,1]=>ARITHM ([Vá.2->Vá.1::56 <= 0])
REIF_278 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_278 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_278.GEQ.REIF_277)])
REIF_279 = [0,1]=>ARITHM ([Vá.2->Vá.1::57 >= 1]), !REIF_279 = [0,1]=>ARITHM ([Vá.2->Vá.1::57 <= 0])
REIF_280 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_280 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_280.GEQ.REIF_279)])
REIF_281 = [0,1]=>ARITHM ([Vá.2->Vá.1::58 >= 1]), !REIF_281 = [0,1]=>ARITHM ([Vá.2->Vá.1::58 <= 0])
REIF_282 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_282 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_282.GEQ.REIF_281)])
REIF_283 = [0,1]=>ARITHM ([Vá.2->Vá.1::59 >= 1]), !REIF_283 = [0,1]=>ARITHM ([Vá.2->Vá.1::59 <= 0])
REIF_284 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_284 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_284.GEQ.REIF_283)])
REIF_285 = [0,1]=>ARITHM ([Vá.2->Vá.1::60 >= 1]), !REIF_285 = [0,1]=>ARITHM ([Vá.2->Vá.1::60 <= 0])
REIF_286 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_286 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_286.GEQ.REIF_285)])
REIF_287 = [0,1]=>ARITHM ([Vá.2->Vá.1::61 >= 1]), !REIF_287 = [0,1]=>ARITHM ([Vá.2->Vá.1::61 <= 0])
REIF_288 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_288 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_288.GEQ.REIF_287)])
REIF_289 = [0,1]=>ARITHM ([Vá.2->Vá.1::62 >= 1]), !REIF_289 = [0,1]=>ARITHM ([Vá.2->Vá.1::62 <= 0])
REIF_290 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_290 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_290.GEQ.REIF_289)])
REIF_291 = [0,1]=>ARITHM ([Vá.2->Vá.1::63 >= 1]), !REIF_291 = [0,1]=>ARITHM ([Vá.2->Vá.1::63 <= 0])
REIF_292 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_292 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_292.GEQ.REIF_291)])
REIF_293 = [0,1]=>ARITHM ([Vá.2->Vá.1::64 >= 1]), !REIF_293 = [0,1]=>ARITHM ([Vá.2->Vá.1::64 <= 0])
REIF_294 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_294 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_294.GEQ.REIF_293)])
REIF_295 = [0,1]=>ARITHM ([Vá.2->Vá.1::65 >= 1]), !REIF_295 = [0,1]=>ARITHM ([Vá.2->Vá.1::65 <= 0])
REIF_296 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_296 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_296.GEQ.REIF_295)])
REIF_297 = [0,1]=>ARITHM ([Vá.2->Vá.1::66 >= 1]), !REIF_297 = [0,1]=>ARITHM ([Vá.2->Vá.1::66 <= 0])
REIF_298 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_298 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_298.GEQ.REIF_297)])
REIF_299 = [0,1]=>ARITHM ([Vá.2->Vá.1::67 >= 1]), !REIF_299 = [0,1]=>ARITHM ([Vá.2->Vá.1::67 <= 0])
REIF_300 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_300 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_300.GEQ.REIF_299)])
REIF_301 = [0,1]=>ARITHM ([Vá.2->Vá.1::68 >= 1]), !REIF_301 = [0,1]=>ARITHM ([Vá.2->Vá.1::68 <= 0])
REIF_302 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_302 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_302.GEQ.REIF_301)])
REIF_303 = [0,1]=>ARITHM ([Vá.2->Vá.1::69 >= 1]), !REIF_303 = [0,1]=>ARITHM ([Vá.2->Vá.1::69 <= 0])
REIF_304 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_304 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_304.GEQ.REIF_303)])
REIF_305 = [0,1]=>ARITHM ([Vá.2->Vá.1::70 >= 1]), !REIF_305 = [0,1]=>ARITHM ([Vá.2->Vá.1::70 <= 0])
REIF_306 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_306 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_306.GEQ.REIF_305)])
REIF_307 = [0,1]=>ARITHM ([Vá.2->Vá.1::71 >= 1]), !REIF_307 = [0,1]=>ARITHM ([Vá.2->Vá.1::71 <= 0])
REIF_308 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_308 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_308.GEQ.REIF_307)])
REIF_309 = [0,1]=>ARITHM ([Vá.2->Vá.1::72 >= 1]), !REIF_309 = [0,1]=>ARITHM ([Vá.2->Vá.1::72 <= 0])
REIF_310 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_310 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_310.GEQ.REIF_309)])
REIF_311 = [0,1]=>ARITHM ([Vá.2->Vá.1::73 >= 1]), !REIF_311 = [0,1]=>ARITHM ([Vá.2->Vá.1::73 <= 0])
REIF_312 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_312 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_312.GEQ.REIF_311)])
REIF_313 = [0,1]=>ARITHM ([Vá.2->Vá.1::74 >= 1]), !REIF_313 = [0,1]=>ARITHM ([Vá.2->Vá.1::74 <= 0])
REIF_314 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_314 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_314.GEQ.REIF_313)])
REIF_315 = [0,1]=>ARITHM ([Vá.2->Vá.1::75 >= 1]), !REIF_315 = [0,1]=>ARITHM ([Vá.2->Vá.1::75 <= 0])
REIF_316 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_316 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_316.GEQ.REIF_315)])
REIF_317 = [0,1]=>ARITHM ([Vá.2->Vá.1::76 >= 1]), !REIF_317 = [0,1]=>ARITHM ([Vá.2->Vá.1::76 <= 0])
REIF_318 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_318 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_318.GEQ.REIF_317)])
REIF_319 = [0,1]=>ARITHM ([Vá.2->Vá.1::77 >= 1]), !REIF_319 = [0,1]=>ARITHM ([Vá.2->Vá.1::77 <= 0])
REIF_320 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_320 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_320.GEQ.REIF_319)])
REIF_321 = [0,1]=>ARITHM ([Vá.2->Vá.1::78 >= 1]), !REIF_321 = [0,1]=>ARITHM ([Vá.2->Vá.1::78 <= 0])
REIF_322 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_322 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_322.GEQ.REIF_321)])
REIF_323 = [0,1]=>ARITHM ([Vá.2->Vá.1::79 >= 1]), !REIF_323 = [0,1]=>ARITHM ([Vá.2->Vá.1::79 <= 0])
REIF_324 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_324 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_324.GEQ.REIF_323)])
REIF_325 = [0,1]=>ARITHM ([Vá.2->Vá.1::80 >= 1]), !REIF_325 = [0,1]=>ARITHM ([Vá.2->Vá.1::80 <= 0])
REIF_326 = [0,1]=>ARITHM ([IV_2 >= IV_1 + -1]), !REIF_326 = [0,1]=>ARITHM ([IV_1 >= IV_2 + 2])
ARITHM ([prop(REIF_326.GEQ.REIF_325)])

Bounds
depot: 0
depot: 999999
Vá.1->Vá.2::0: 328
Vá.1->Vá.2::0: 360
Vá.1->Vá.2::1: 348
Vá.1->Vá.2::1: 380
Vá.1->Vá.2::2: 368
Vá.1->Vá.2::2: 400
Vá.1->Vá.2::3: 388
Vá.1->Vá.2::3: 421
Vá.1->Vá.2::4: 408
Vá.1->Vá.2::4: 441
Vá.1->Vá.2::5: 418
Vá.1->Vá.2::5: 452
Vá.1->Vá.2::6: 428
Vá.1->Vá.2::6: 462
Vá.1->Vá.2::7: 438
Vá.1->Vá.2::7: 473
Vá.1->Vá.2::8: 451
Vá.1->Vá.2::8: 487
Vá.1->Vá.2::9: 461
Vá.1->Vá.2::9: 497
Vá.1->Vá.2::10: 471
Vá.1->Vá.2::10: 508
Vá.1->Vá.2::11: 481
Vá.1->Vá.2::11: 519
Vá.1->Vá.2::12: 491
Vá.1->Vá.2::12: 529
Vá.1->Vá.2::13: 501
Vá.1->Vá.2::13: 540
Vá.1->Vá.2::14: 511
Vá.1->Vá.2::14: 551
Vá.1->Vá.2::15: 521
Vá.1->Vá.2::15: 561
Vá.1->Vá.2::16: 531
Vá.1->Vá.2::16: 571
Vá.1->Vá.2::17: 541
Vá.1->Vá.2::17: 580
Vá.1->Vá.2::18: 551
Vá.1->Vá.2::18: 590
Vá.1->Vá.2::19: 561
Vá.1->Vá.2::19: 600
Vá.1->Vá.2::20: 571
Vá.1->Vá.2::20: 609
Vá.1->Vá.2::21: 581
Vá.1->Vá.2::21: 619
Vá.1->Vá.2::22: 593
Vá.1->Vá.2::22: 631
Vá.1->Vá.2::23: 608
Vá.1->Vá.2::23: 645
Vá.1->Vá.2::24: 623
Vá.1->Vá.2::24: 660
Vá.1->Vá.2::25: 638
Vá.1->Vá.2::25: 674
Vá.1->Vá.2::26: 653
Vá.1->Vá.2::26: 689
Vá.1->Vá.2::27: 668
Vá.1->Vá.2::27: 704
Vá.1->Vá.2::28: 683
Vá.1->Vá.2::28: 719
Vá.1->Vá.2::29: 698
Vá.1->Vá.2::29: 734
Vá.1->Vá.2::30: 713
Vá.1->Vá.2::30: 749
Vá.1->Vá.2::31: 728
Vá.1->Vá.2::31: 764
Vá.1->Vá.2::32: 743
Vá.1->Vá.2::32: 779
Vá.1->Vá.2::33: 758
Vá.1->Vá.2::33: 794
Vá.1->Vá.2::34: 773
Vá.1->Vá.2::34: 809
Vá.1->Vá.2::35: 788
Vá.1->Vá.2::35: 824
Vá.1->Vá.2::36: 803
Vá.1->Vá.2::36: 839
Vá.1->Vá.2::37: 818
Vá.1->Vá.2::37: 854
Vá.1->Vá.2::38: 833
Vá.1->Vá.2::38: 869
Vá.1->Vá.2::39: 848
Vá.1->Vá.2::39: 884
Vá.1->Vá.2::40: 863
Vá.1->Vá.2::40: 899
Vá.1->Vá.2::41: 878
Vá.1->Vá.2::41: 914
Vá.1->Vá.2::42: 893
Vá.1->Vá.2::42: 929
Vá.1->Vá.2::43: 908
Vá.1->Vá.2::43: 944
Vá.1->Vá.2::44: 923
Vá.1->Vá.2::44: 959
Vá.1->Vá.2::45: 935
Vá.1->Vá.2::45: 971
Vá.1->Vá.2::46: 947
Vá.1->Vá.2::46: 983
Vá.1->Vá.2::47: 959
Vá.1->Vá.2::47: 995
Vá.1->Vá.2::48: 971
Vá.1->Vá.2::48: 1007
Vá.1->Vá.2::49: 983
Vá.1->Vá.2::49: 1019
Vá.1->Vá.2::50: 995
Vá.1->Vá.2::50: 1032
Vá.1->Vá.2::51: 1007
Vá.1->Vá.2::51: 1044
Vá.1->Vá.2::52: 1019
Vá.1->Vá.2::52: 1057
Vá.1->Vá.2::53: 1031
Vá.1->Vá.2::53: 1069
Vá.1->Vá.2::54: 1043
Vá.1->Vá.2::54: 1081
Vá.1->Vá.2::55: 1055
Vá.1->Vá.2::55: 1094
Vá.1->Vá.2::56: 1067
Vá.1->Vá.2::56: 1106
Vá.1->Vá.2::57: 1079
Vá.1->Vá.2::57: 1118
Vá.1->Vá.2::58: 1091
Vá.1->Vá.2::58: 1130
Vá.1->Vá.2::59: 1103
Vá.1->Vá.2::59: 1142
Vá.1->Vá.2::60: 1115
Vá.1->Vá.2::60: 1153
Vá.1->Vá.2::61: 1127
Vá.1->Vá.2::61: 1165
Vá.1->Vá.2::62: 1139
Vá.1->Vá.2::62: 1176
Vá.1->Vá.2::63: 1151
Vá.1->Vá.2::63: 1188
Vá.1->Vá.2::64: 1163
Vá.1->Vá.2::64: 1200
Vá.1->Vá.2::65: 1172
Vá.1->Vá.2::65: 1208
Vá.1->Vá.2::66: 1184
Vá.1->Vá.2::66: 1220
Vá.1->Vá.2::67: 1197
Vá.1->Vá.2::67: 1232
Vá.1->Vá.2::68: 1211
Vá.1->Vá.2::68: 1246
Vá.1->Vá.2::69: 1224
Vá.1->Vá.2::69: 1259
Vá.1->Vá.2::70: 1238
Vá.1->Vá.2::70: 1272
Vá.1->Vá.2::71: 1253
Vá.1->Vá.2::71: 1287
Vá.1->Vá.2::72: 1268
Vá.1->Vá.2::72: 1301
Vá.1->Vá.2::73: 1283
Vá.1->Vá.2::73: 1316
Vá.1->Vá.2::74: 1298
Vá.1->Vá.2::74: 1330
Vá.1->Vá.2::75: 1313
Vá.1->Vá.2::75: 1345
Vá.1->Vá.2::76: 1328
Vá.1->Vá.2::76: 1359
Vá.1->Vá.2::77: 1343
Vá.1->Vá.2::77: 1374
Vá.1->Vá.2::78: 1358
Vá.1->Vá.2::78: 1387
Vá.1->Vá.2::79: 1378
Vá.1->Vá.2::79: 1407
Vá.1->Vá.2::80: 1398
Vá.1->Vá.2::80: 1427
Vá.2->Vá.1::0: 262
Vá.2->Vá.1::0: 293
Vá.2->Vá.1::1: 282
Vá.2->Vá.1::1: 313
Vá.2->Vá.1::2: 302
Vá.2->Vá.1::2: 335
Vá.2->Vá.1::3: 322
Vá.2->Vá.1::3: 355
Vá.2->Vá.1::4: 332
Vá.2->Vá.1::4: 366
Vá.2->Vá.1::5: 342
Vá.2->Vá.1::5: 376
Vá.2->Vá.1::6: 352
Vá.2->Vá.1::6: 386
Vá.2->Vá.1::7: 362
Vá.2->Vá.1::7: 397
Vá.2->Vá.1::8: 372
Vá.2->Vá.1::8: 407
Vá.2->Vá.1::9: 382
Vá.2->Vá.1::9: 418
Vá.2->Vá.1::10: 392
Vá.2->Vá.1::10: 429
Vá.2->Vá.1::11: 402
Vá.2->Vá.1::11: 439
Vá.2->Vá.1::12: 412
Vá.2->Vá.1::12: 450
Vá.2->Vá.1::13: 422
Vá.2->Vá.1::13: 461
Vá.2->Vá.1::14: 432
Vá.2->Vá.1::14: 471
Vá.2->Vá.1::15: 442
Vá.2->Vá.1::15: 483
Vá.2->Vá.1::16: 452
Vá.2->Vá.1::16: 494
Vá.2->Vá.1::17: 462
Vá.2->Vá.1::17: 504
Vá.2->Vá.1::18: 472
Vá.2->Vá.1::18: 514
Vá.2->Vá.1::19: 482
Vá.2->Vá.1::19: 523
Vá.2->Vá.1::20: 492
Vá.2->Vá.1::20: 533
Vá.2->Vá.1::21: 502
Vá.2->Vá.1::21: 543
Vá.2->Vá.1::22: 517
Vá.2->Vá.1::22: 558
Vá.2->Vá.1::23: 532
Vá.2->Vá.1::23: 573
Vá.2->Vá.1::24: 547
Vá.2->Vá.1::24: 587
Vá.2->Vá.1::25: 562
Vá.2->Vá.1::25: 602
Vá.2->Vá.1::26: 577
Vá.2->Vá.1::26: 617
Vá.2->Vá.1::27: 592
Vá.2->Vá.1::27: 632
Vá.2->Vá.1::28: 607
Vá.2->Vá.1::28: 646
Vá.2->Vá.1::29: 622
Vá.2->Vá.1::29: 661
Vá.2->Vá.1::30: 637
Vá.2->Vá.1::30: 676
Vá.2->Vá.1::31: 652
Vá.2->Vá.1::31: 691
Vá.2->Vá.1::32: 667
Vá.2->Vá.1::32: 706
Vá.2->Vá.1::33: 682
Vá.2->Vá.1::33: 721
Vá.2->Vá.1::34: 697
Vá.2->Vá.1::34: 736
Vá.2->Vá.1::35: 712
Vá.2->Vá.1::35: 751
Vá.2->Vá.1::36: 727
Vá.2->Vá.1::36: 766
Vá.2->Vá.1::37: 742
Vá.2->Vá.1::37: 781
Vá.2->Vá.1::38: 757
Vá.2->Vá.1::38: 796
Vá.2->Vá.1::39: 772
Vá.2->Vá.1::39: 811
Vá.2->Vá.1::40: 787
Vá.2->Vá.1::40: 826
Vá.2->Vá.1::41: 802
Vá.2->Vá.1::41: 841
Vá.2->Vá.1::42: 817
Vá.2->Vá.1::42: 856
Vá.2->Vá.1::43: 832
Vá.2->Vá.1::43: 871
Vá.2->Vá.1::44: 847
Vá.2->Vá.1::44: 886
Vá.2->Vá.1::45: 859
Vá.2->Vá.1::45: 898
Vá.2->Vá.1::46: 871
Vá.2->Vá.1::46: 910
Vá.2->Vá.1::47: 883
Vá.2->Vá.1::47: 922
Vá.2->Vá.1::48: 895
Vá.2->Vá.1::48: 934
Vá.2->Vá.1::49: 907
Vá.2->Vá.1::49: 946
Vá.2->Vá.1::50: 919
Vá.2->Vá.1::50: 958
Vá.2->Vá.1::51: 931
Vá.2->Vá.1::51: 970
Vá.2->Vá.1::52: 943
Vá.2->Vá.1::52: 982
Vá.2->Vá.1::53: 955
Vá.2->Vá.1::53: 994
Vá.2->Vá.1::54: 967
Vá.2->Vá.1::54: 1007
Vá.2->Vá.1::55: 979
Vá.2->Vá.1::55: 1019
Vá.2->Vá.1::56: 991
Vá.2->Vá.1::56: 1031
Vá.2->Vá.1::57: 1003
Vá.2->Vá.1::57: 1043
Vá.2->Vá.1::58: 1015
Vá.2->Vá.1::58: 1055
Vá.2->Vá.1::59: 1027
Vá.2->Vá.1::59: 1067
Vá.2->Vá.1::60: 1039
Vá.2->Vá.1::60: 1079
Vá.2->Vá.1::61: 1051
Vá.2->Vá.1::61: 1091
Vá.2->Vá.1::62: 1063
Vá.2->Vá.1::62: 1103
Vá.2->Vá.1::63: 1075
Vá.2->Vá.1::63: 1115
Vá.2->Vá.1::64: 1087
Vá.2->Vá.1::64: 1126
Vá.2->Vá.1::65: 1099
Vá.2->Vá.1::65: 1137
Vá.2->Vá.1::66: 1111
Vá.2->Vá.1::66: 1148
Vá.2->Vá.1::67: 1125
Vá.2->Vá.1::67: 1161
Vá.2->Vá.1::68: 1140
Vá.2->Vá.1::68: 1175
Vá.2->Vá.1::69: 1155
Vá.2->Vá.1::69: 1190
Vá.2->Vá.1::70: 1170
Vá.2->Vá.1::70: 1204
Vá.2->Vá.1::71: 1185
Vá.2->Vá.1::71: 1219
Vá.2->Vá.1::72: 1200
Vá.2->Vá.1::72: 1233
Vá.2->Vá.1::73: 1215
Vá.2->Vá.1::73: 1248
Vá.2->Vá.1::74: 1230
Vá.2->Vá.1::74: 1262
Vá.2->Vá.1::75: 1245
Vá.2->Vá.1::75: 1277
Vá.2->Vá.1::76: 1260
Vá.2->Vá.1::76: 1292
Vá.2->Vá.1::77: 1278
Vá.2->Vá.1::77: 1310
Vá.2->Vá.1::78: 1297
Vá.2->Vá.1::78: 1327
Vá.2->Vá.1::79: 1317
Vá.2->Vá.1::79: 1347
Vá.2->Vá.1::80: 1337
Vá.2->Vá.1::80: 1367
IV_1: 1
IV_1: 162
IV_2: 0
IV_2: 162
cste -- 1: 1
cste -- 1: 1
REIF_3: 0
REIF_3: 1
not(REIF_3): 0
not(REIF_3): 1
REIF_4: 0
REIF_4: 1
not(REIF_4): 0
not(REIF_4): 1
REIF_5: 0
REIF_5: 1
not(REIF_5): 0
not(REIF_5): 1
REIF_6: 0
REIF_6: 1
not(REIF_6): 0
not(REIF_6): 1
REIF_7: 0
REIF_7: 1
not(REIF_7): 0
not(REIF_7): 1
REIF_8: 0
REIF_8: 1
not(REIF_8): 0
not(REIF_8): 1
REIF_9: 0
REIF_9: 1
not(REIF_9): 0
not(REIF_9): 1
REIF_10: 0
REIF_10: 1
not(REIF_10): 0
not(REIF_10): 1
REIF_11: 0
REIF_11: 1
not(REIF_11): 0
not(REIF_11): 1
REIF_12: 0
REIF_12: 1
not(REIF_12): 0
not(REIF_12): 1
REIF_13: 0
REIF_13: 1
not(REIF_13): 0
not(REIF_13): 1
REIF_14: 0
REIF_14: 1
not(REIF_14): 0
not(REIF_14): 1
REIF_15: 0
REIF_15: 1
not(REIF_15): 0
not(REIF_15): 1
REIF_16: 0
REIF_16: 1
not(REIF_16): 0
not(REIF_16): 1
REIF_17: 0
REIF_17: 1
not(REIF_17): 0
not(REIF_17): 1
REIF_18: 0
REIF_18: 1
not(REIF_18): 0
not(REIF_18): 1
REIF_19: 0
REIF_19: 1
not(REIF_19): 0
not(REIF_19): 1
REIF_20: 0
REIF_20: 1
not(REIF_20): 0
not(REIF_20): 1
REIF_21: 0
REIF_21: 1
not(REIF_21): 0
not(REIF_21): 1
REIF_22: 0
REIF_22: 1
not(REIF_22): 0
not(REIF_22): 1
REIF_23: 0
REIF_23: 1
not(REIF_23): 0
not(REIF_23): 1
REIF_24: 0
REIF_24: 1
not(REIF_24): 0
not(REIF_24): 1
REIF_25: 0
REIF_25: 1
not(REIF_25): 0
not(REIF_25): 1
REIF_26: 0
REIF_26: 1
not(REIF_26): 0
not(REIF_26): 1
REIF_27: 0
REIF_27: 1
not(REIF_27): 0
not(REIF_27): 1
REIF_28: 0
REIF_28: 1
not(REIF_28): 0
not(REIF_28): 1
REIF_29: 0
REIF_29: 1
not(REIF_29): 0
not(REIF_29): 1
REIF_30: 0
REIF_30: 1
not(REIF_30): 0
not(REIF_30): 1
REIF_31: 0
REIF_31: 1
not(REIF_31): 0
not(REIF_31): 1
REIF_32: 0
REIF_32: 1
not(REIF_32): 0
not(REIF_32): 1
REIF_33: 0
REIF_33: 1
not(REIF_33): 0
not(REIF_33): 1
REIF_34: 0
REIF_34: 1
not(REIF_34): 0
not(REIF_34): 1
REIF_35: 0
REIF_35: 1
not(REIF_35): 0
not(REIF_35): 1
REIF_36: 0
REIF_36: 1
not(REIF_36): 0
not(REIF_36): 1
REIF_37: 0
REIF_37: 1
not(REIF_37): 0
not(REIF_37): 1
REIF_38: 0
REIF_38: 1
not(REIF_38): 0
not(REIF_38): 1
REIF_39: 0
REIF_39: 1
not(REIF_39): 0
not(REIF_39): 1
REIF_40: 0
REIF_40: 1
not(REIF_40): 0
not(REIF_40): 1
REIF_41: 0
REIF_41: 1
not(REIF_41): 0
not(REIF_41): 1
REIF_42: 0
REIF_42: 1
not(REIF_42): 0
not(REIF_42): 1
REIF_43: 0
REIF_43: 1
not(REIF_43): 0
not(REIF_43): 1
REIF_44: 0
REIF_44: 1
not(REIF_44): 0
not(REIF_44): 1
REIF_45: 0
REIF_45: 1
not(REIF_45): 0
not(REIF_45): 1
REIF_46: 0
REIF_46: 1
not(REIF_46): 0
not(REIF_46): 1
REIF_47: 0
REIF_47: 1
not(REIF_47): 0
not(REIF_47): 1
REIF_48: 0
REIF_48: 1
not(REIF_48): 0
not(REIF_48): 1
REIF_49: 0
REIF_49: 1
not(REIF_49): 0
not(REIF_49): 1
REIF_50: 0
REIF_50: 1
not(REIF_50): 0
not(REIF_50): 1
REIF_51: 0
REIF_51: 1
not(REIF_51): 0
not(REIF_51): 1
REIF_52: 0
REIF_52: 1
not(REIF_52): 0
not(REIF_52): 1
REIF_53: 0
REIF_53: 1
not(REIF_53): 0
not(REIF_53): 1
REIF_54: 0
REIF_54: 1
not(REIF_54): 0
not(REIF_54): 1
REIF_55: 0
REIF_55: 1
not(REIF_55): 0
not(REIF_55): 1
REIF_56: 0
REIF_56: 1
not(REIF_56): 0
not(REIF_56): 1
REIF_57: 0
REIF_57: 1
not(REIF_57): 0
not(REIF_57): 1
REIF_58: 0
REIF_58: 1
not(REIF_58): 0
not(REIF_58): 1
REIF_59: 0
REIF_59: 1
not(REIF_59): 0
not(REIF_59): 1
REIF_60: 0
REIF_60: 1
not(REIF_60): 0
not(REIF_60): 1
REIF_61: 0
REIF_61: 1
not(REIF_61): 0
not(REIF_61): 1
REIF_62: 0
REIF_62: 1
not(REIF_62): 0
not(REIF_62): 1
REIF_63: 0
REIF_63: 1
not(REIF_63): 0
not(REIF_63): 1
REIF_64: 0
REIF_64: 1
not(REIF_64): 0
not(REIF_64): 1
REIF_65: 0
REIF_65: 1
not(REIF_65): 0
not(REIF_65): 1
REIF_66: 0
REIF_66: 1
not(REIF_66): 0
not(REIF_66): 1
REIF_67: 0
REIF_67: 1
not(REIF_67): 0
not(REIF_67): 1
REIF_68: 0
REIF_68: 1
not(REIF_68): 0
not(REIF_68): 1
REIF_69: 0
REIF_69: 1
not(REIF_69): 0
not(REIF_69): 1
REIF_70: 0
REIF_70: 1
not(REIF_70): 0
not(REIF_70): 1
REIF_71: 0
REIF_71: 1
not(REIF_71): 0
not(REIF_71): 1
REIF_72: 0
REIF_72: 1
not(REIF_72): 0
not(REIF_72): 1
REIF_73: 0
REIF_73: 1
not(REIF_73): 0
not(REIF_73): 1
REIF_74: 0
REIF_74: 1
not(REIF_74): 0
not(REIF_74): 1
REIF_75: 0
REIF_75: 1
not(REIF_75): 0
not(REIF_75): 1
REIF_76: 0
REIF_76: 1
not(REIF_76): 0
not(REIF_76): 1
REIF_77: 0
REIF_77: 1
not(REIF_77): 0
not(REIF_77): 1
REIF_78: 0
REIF_78: 1
not(REIF_78): 0
not(REIF_78): 1
REIF_79: 0
REIF_79: 1
not(REIF_79): 0
not(REIF_79): 1
REIF_80: 0
REIF_80: 1
not(REIF_80): 0
not(REIF_80): 1
REIF_81: 0
REIF_81: 1
not(REIF_81): 0
not(REIF_81): 1
REIF_82: 0
REIF_82: 1
not(REIF_82): 0
not(REIF_82): 1
REIF_83: 0
REIF_83: 1
not(REIF_83): 0
not(REIF_83): 1
REIF_84: 0
REIF_84: 1
not(REIF_84): 0
not(REIF_84): 1
REIF_85: 0
REIF_85: 1
not(REIF_85): 0
not(REIF_85): 1
REIF_86: 0
REIF_86: 1
not(REIF_86): 0
not(REIF_86): 1
REIF_87: 0
REIF_87: 1
not(REIF_87): 0
not(REIF_87): 1
REIF_88: 0
REIF_88: 1
not(REIF_88): 0
not(REIF_88): 1
REIF_89: 0
REIF_89: 1
not(REIF_89): 0
not(REIF_89): 1
REIF_90: 0
REIF_90: 1
not(REIF_90): 0
not(REIF_90): 1
REIF_91: 0
REIF_91: 1
not(REIF_91): 0
not(REIF_91): 1
REIF_92: 0
REIF_92: 1
not(REIF_92): 0
not(REIF_92): 1
REIF_93: 0
REIF_93: 1
not(REIF_93): 0
not(REIF_93): 1
REIF_94: 0
REIF_94: 1
not(REIF_94): 0
not(REIF_94): 1
REIF_95: 0
REIF_95: 1
not(REIF_95): 0
not(REIF_95): 1
REIF_96: 0
REIF_96: 1
not(REIF_96): 0
not(REIF_96): 1
REIF_97: 0
REIF_97: 1
not(REIF_97): 0
not(REIF_97): 1
REIF_98: 0
REIF_98: 1
not(REIF_98): 0
not(REIF_98): 1
REIF_99: 0
REIF_99: 1
not(REIF_99): 0
not(REIF_99): 1
REIF_100: 0
REIF_100: 1
not(REIF_100): 0
not(REIF_100): 1
REIF_101: 0
REIF_101: 1
not(REIF_101): 0
not(REIF_101): 1
REIF_102: 0
REIF_102: 1
not(REIF_102): 0
not(REIF_102): 1
REIF_103: 0
REIF_103: 1
not(REIF_103): 0
not(REIF_103): 1
REIF_104: 0
REIF_104: 1
not(REIF_104): 0
not(REIF_104): 1
REIF_105: 0
REIF_105: 1
not(REIF_105): 0
not(REIF_105): 1
REIF_106: 0
REIF_106: 1
not(REIF_106): 0
not(REIF_106): 1
REIF_107: 0
REIF_107: 1
not(REIF_107): 0
not(REIF_107): 1
REIF_108: 0
REIF_108: 1
not(REIF_108): 0
not(REIF_108): 1
REIF_109: 0
REIF_109: 1
not(REIF_109): 0
not(REIF_109): 1
REIF_110: 0
REIF_110: 1
not(REIF_110): 0
not(REIF_110): 1
REIF_111: 0
REIF_111: 1
not(REIF_111): 0
not(REIF_111): 1
REIF_112: 0
REIF_112: 1
not(REIF_112): 0
not(REIF_112): 1
REIF_113: 0
REIF_113: 1
not(REIF_113): 0
not(REIF_113): 1
REIF_114: 0
REIF_114: 1
not(REIF_114): 0
not(REIF_114): 1
REIF_115: 0
REIF_115: 1
not(REIF_115): 0
not(REIF_115): 1
REIF_116: 0
REIF_116: 1
not(REIF_116): 0
not(REIF_116): 1
REIF_117: 0
REIF_117: 1
not(REIF_117): 0
not(REIF_117): 1
REIF_118: 0
REIF_118: 1
not(REIF_118): 0
not(REIF_118): 1
REIF_119: 0
REIF_119: 1
not(REIF_119): 0
not(REIF_119): 1
REIF_120: 0
REIF_120: 1
not(REIF_120): 0
not(REIF_120): 1
REIF_121: 0
REIF_121: 1
not(REIF_121): 0
not(REIF_121): 1
REIF_122: 0
REIF_122: 1
not(REIF_122): 0
not(REIF_122): 1
REIF_123: 0
REIF_123: 1
not(REIF_123): 0
not(REIF_123): 1
REIF_124: 0
REIF_124: 1
not(REIF_124): 0
not(REIF_124): 1
REIF_125: 0
REIF_125: 1
not(REIF_125): 0
not(REIF_125): 1
REIF_126: 0
REIF_126: 1
not(REIF_126): 0
not(REIF_126): 1
REIF_127: 0
REIF_127: 1
not(REIF_127): 0
not(REIF_127): 1
REIF_128: 0
REIF_128: 1
not(REIF_128): 0
not(REIF_128): 1
REIF_129: 0
REIF_129: 1
not(REIF_129): 0
not(REIF_129): 1
REIF_130: 0
REIF_130: 1
not(REIF_130): 0
not(REIF_130): 1
REIF_131: 0
REIF_131: 1
not(REIF_131): 0
not(REIF_131): 1
REIF_132: 0
REIF_132: 1
not(REIF_132): 0
not(REIF_132): 1
REIF_133: 0
REIF_133: 1
not(REIF_133): 0
not(REIF_133): 1
REIF_134: 0
REIF_134: 1
not(REIF_134): 0
not(REIF_134): 1
REIF_135: 0
REIF_135: 1
not(REIF_135): 0
not(REIF_135): 1
REIF_136: 0
REIF_136: 1
not(REIF_136): 0
not(REIF_136): 1
REIF_137: 0
REIF_137: 1
not(REIF_137): 0
not(REIF_137): 1
REIF_138: 0
REIF_138: 1
not(REIF_138): 0
not(REIF_138): 1
REIF_139: 0
REIF_139: 1
not(REIF_139): 0
not(REIF_139): 1
REIF_140: 0
REIF_140: 1
not(REIF_140): 0
not(REIF_140): 1
REIF_141: 0
REIF_141: 1
not(REIF_141): 0
not(REIF_141): 1
REIF_142: 0
REIF_142: 1
not(REIF_142): 0
not(REIF_142): 1
REIF_143: 0
REIF_143: 1
not(REIF_143): 0
not(REIF_143): 1
REIF_144: 0
REIF_144: 1
not(REIF_144): 0
not(REIF_144): 1
REIF_145: 0
REIF_145: 1
not(REIF_145): 0
not(REIF_145): 1
REIF_146: 0
REIF_146: 1
not(REIF_146): 0
not(REIF_146): 1
REIF_147: 0
REIF_147: 1
not(REIF_147): 0
not(REIF_147): 1
REIF_148: 0
REIF_148: 1
not(REIF_148): 0
not(REIF_148): 1
REIF_149: 0
REIF_149: 1
not(REIF_149): 0
not(REIF_149): 1
REIF_150: 0
REIF_150: 1
not(REIF_150): 0
not(REIF_150): 1
REIF_151: 0
REIF_151: 1
not(REIF_151): 0
not(REIF_151): 1
REIF_152: 0
REIF_152: 1
not(REIF_152): 0
not(REIF_152): 1
REIF_153: 0
REIF_153: 1
not(REIF_153): 0
not(REIF_153): 1
REIF_154: 0
REIF_154: 1
not(REIF_154): 0
not(REIF_154): 1
REIF_155: 0
REIF_155: 1
not(REIF_155): 0
not(REIF_155): 1
REIF_156: 0
REIF_156: 1
not(REIF_156): 0
not(REIF_156): 1
REIF_157: 0
REIF_157: 1
not(REIF_157): 0
not(REIF_157): 1
REIF_158: 0
REIF_158: 1
not(REIF_158): 0
not(REIF_158): 1
REIF_159: 0
REIF_159: 1
not(REIF_159): 0
not(REIF_159): 1
REIF_160: 0
REIF_160: 1
not(REIF_160): 0
not(REIF_160): 1
REIF_161: 0
REIF_161: 1
not(REIF_161): 0
not(REIF_161): 1
REIF_162: 0
REIF_162: 1
not(REIF_162): 0
not(REIF_162): 1
REIF_163: 0
REIF_163: 1
not(REIF_163): 0
not(REIF_163): 1
REIF_164: 0
REIF_164: 1
not(REIF_164): 0
not(REIF_164): 1
REIF_165: 0
REIF_165: 1
not(REIF_165): 0
not(REIF_165): 1
REIF_166: 0
REIF_166: 1
not(REIF_166): 0
not(REIF_166): 1
REIF_167: 0
REIF_167: 1
not(REIF_167): 0
not(REIF_167): 1
REIF_168: 0
REIF_168: 1
not(REIF_168): 0
not(REIF_168): 1
REIF_169: 0
REIF_169: 1
not(REIF_169): 0
not(REIF_169): 1
REIF_170: 0
REIF_170: 1
not(REIF_170): 0
not(REIF_170): 1
REIF_171: 0
REIF_171: 1
not(REIF_171): 0
not(REIF_171): 1
REIF_172: 0
REIF_172: 1
not(REIF_172): 0
not(REIF_172): 1
REIF_173: 0
REIF_173: 1
not(REIF_173): 0
not(REIF_173): 1
REIF_174: 0
REIF_174: 1
not(REIF_174): 0
not(REIF_174): 1
REIF_175: 0
REIF_175: 1
not(REIF_175): 0
not(REIF_175): 1
REIF_176: 0
REIF_176: 1
not(REIF_176): 0
not(REIF_176): 1
REIF_177: 0
REIF_177: 1
not(REIF_177): 0
not(REIF_177): 1
REIF_178: 0
REIF_178: 1
not(REIF_178): 0
not(REIF_178): 1
REIF_179: 0
REIF_179: 1
not(REIF_179): 0
not(REIF_179): 1
REIF_180: 0
REIF_180: 1
not(REIF_180): 0
not(REIF_180): 1
REIF_181: 0
REIF_181: 1
not(REIF_181): 0
not(REIF_181): 1
REIF_182: 0
REIF_182: 1
not(REIF_182): 0
not(REIF_182): 1
REIF_183: 0
REIF_183: 1
not(REIF_183): 0
not(REIF_183): 1
REIF_184: 0
REIF_184: 1
not(REIF_184): 0
not(REIF_184): 1
REIF_185: 0
REIF_185: 1
not(REIF_185): 0
not(REIF_185): 1
REIF_186: 0
REIF_186: 1
not(REIF_186): 0
not(REIF_186): 1
REIF_187: 0
REIF_187: 1
not(REIF_187): 0
not(REIF_187): 1
REIF_188: 0
REIF_188: 1
not(REIF_188): 0
not(REIF_188): 1
REIF_189: 0
REIF_189: 1
not(REIF_189): 0
not(REIF_189): 1
REIF_190: 0
REIF_190: 1
not(REIF_190): 0
not(REIF_190): 1
REIF_191: 0
REIF_191: 1
not(REIF_191): 0
not(REIF_191): 1
REIF_192: 0
REIF_192: 1
not(REIF_192): 0
not(REIF_192): 1
REIF_193: 0
REIF_193: 1
not(REIF_193): 0
not(REIF_193): 1
REIF_194: 0
REIF_194: 1
not(REIF_194): 0
not(REIF_194): 1
REIF_195: 0
REIF_195: 1
not(REIF_195): 0
not(REIF_195): 1
REIF_196: 0
REIF_196: 1
not(REIF_196): 0
not(REIF_196): 1
REIF_197: 0
REIF_197: 1
not(REIF_197): 0
not(REIF_197): 1
REIF_198: 0
REIF_198: 1
not(REIF_198): 0
not(REIF_198): 1
REIF_199: 0
REIF_199: 1
not(REIF_199): 0
not(REIF_199): 1
REIF_200: 0
REIF_200: 1
not(REIF_200): 0
not(REIF_200): 1
REIF_201: 0
REIF_201: 1
not(REIF_201): 0
not(REIF_201): 1
REIF_202: 0
REIF_202: 1
not(REIF_202): 0
not(REIF_202): 1
REIF_203: 0
REIF_203: 1
not(REIF_203): 0
not(REIF_203): 1
REIF_204: 0
REIF_204: 1
not(REIF_204): 0
not(REIF_204): 1
REIF_205: 0
REIF_205: 1
not(REIF_205): 0
not(REIF_205): 1
REIF_206: 0
REIF_206: 1
not(REIF_206): 0
not(REIF_206): 1
REIF_207: 0
REIF_207: 1
not(REIF_207): 0
not(REIF_207): 1
REIF_208: 0
REIF_208: 1
not(REIF_208): 0
not(REIF_208): 1
REIF_209: 0
REIF_209: 1
not(REIF_209): 0
not(REIF_209): 1
REIF_210: 0
REIF_210: 1
not(REIF_210): 0
not(REIF_210): 1
REIF_211: 0
REIF_211: 1
not(REIF_211): 0
not(REIF_211): 1
REIF_212: 0
REIF_212: 1
not(REIF_212): 0
not(REIF_212): 1
REIF_213: 0
REIF_213: 1
not(REIF_213): 0
not(REIF_213): 1
REIF_214: 0
REIF_214: 1
not(REIF_214): 0
not(REIF_214): 1
REIF_215: 0
REIF_215: 1
not(REIF_215): 0
not(REIF_215): 1
REIF_216: 0
REIF_216: 1
not(REIF_216): 0
not(REIF_216): 1
REIF_217: 0
REIF_217: 1
not(REIF_217): 0
not(REIF_217): 1
REIF_218: 0
REIF_218: 1
not(REIF_218): 0
not(REIF_218): 1
REIF_219: 0
REIF_219: 1
not(REIF_219): 0
not(REIF_219): 1
REIF_220: 0
REIF_220: 1
not(REIF_220): 0
not(REIF_220): 1
REIF_221: 0
REIF_221: 1
not(REIF_221): 0
not(REIF_221): 1
REIF_222: 0
REIF_222: 1
not(REIF_222): 0
not(REIF_222): 1
REIF_223: 0
REIF_223: 1
not(REIF_223): 0
not(REIF_223): 1
REIF_224: 0
REIF_224: 1
not(REIF_224): 0
not(REIF_224): 1
REIF_225: 0
REIF_225: 1
not(REIF_225): 0
not(REIF_225): 1
REIF_226: 0
REIF_226: 1
not(REIF_226): 0
not(REIF_226): 1
REIF_227: 0
REIF_227: 1
not(REIF_227): 0
not(REIF_227): 1
REIF_228: 0
REIF_228: 1
not(REIF_228): 0
not(REIF_228): 1
REIF_229: 0
REIF_229: 1
not(REIF_229): 0
not(REIF_229): 1
REIF_230: 0
REIF_230: 1
not(REIF_230): 0
not(REIF_230): 1
REIF_231: 0
REIF_231: 1
not(REIF_231): 0
not(REIF_231): 1
REIF_232: 0
REIF_232: 1
not(REIF_232): 0
not(REIF_232): 1
REIF_233: 0
REIF_233: 1
not(REIF_233): 0
not(REIF_233): 1
REIF_234: 0
REIF_234: 1
not(REIF_234): 0
not(REIF_234): 1
REIF_235: 0
REIF_235: 1
not(REIF_235): 0
not(REIF_235): 1
REIF_236: 0
REIF_236: 1
not(REIF_236): 0
not(REIF_236): 1
REIF_237: 0
REIF_237: 1
not(REIF_237): 0
not(REIF_237): 1
REIF_238: 0
REIF_238: 1
not(REIF_238): 0
not(REIF_238): 1
REIF_239: 0
REIF_239: 1
not(REIF_239): 0
not(REIF_239): 1
REIF_240: 0
REIF_240: 1
not(REIF_240): 0
not(REIF_240): 1
REIF_241: 0
REIF_241: 1
not(REIF_241): 0
not(REIF_241): 1
REIF_242: 0
REIF_242: 1
not(REIF_242): 0
not(REIF_242): 1
REIF_243: 0
REIF_243: 1
not(REIF_243): 0
not(REIF_243): 1
REIF_244: 0
REIF_244: 1
not(REIF_244): 0
not(REIF_244): 1
REIF_245: 0
REIF_245: 1
not(REIF_245): 0
not(REIF_245): 1
REIF_246: 0
REIF_246: 1
not(REIF_246): 0
not(REIF_246): 1
REIF_247: 0
REIF_247: 1
not(REIF_247): 0
not(REIF_247): 1
REIF_248: 0
REIF_248: 1
not(REIF_248): 0
not(REIF_248): 1
REIF_249: 0
REIF_249: 1
not(REIF_249): 0
not(REIF_249): 1
REIF_250: 0
REIF_250: 1
not(REIF_250): 0
not(REIF_250): 1
REIF_251: 0
REIF_251: 1
not(REIF_251): 0
not(REIF_251): 1
REIF_252: 0
REIF_252: 1
not(REIF_252): 0
not(REIF_252): 1
REIF_253: 0
REIF_253: 1
not(REIF_253): 0
not(REIF_253): 1
REIF_254: 0
REIF_254: 1
not(REIF_254): 0
not(REIF_254): 1
REIF_255: 0
REIF_255: 1
not(REIF_255): 0
not(REIF_255): 1
REIF_256: 0
REIF_256: 1
not(REIF_256): 0
not(REIF_256): 1
REIF_257: 0
REIF_257: 1
not(REIF_257): 0
not(REIF_257): 1
REIF_258: 0
REIF_258: 1
not(REIF_258): 0
not(REIF_258): 1
REIF_259: 0
REIF_259: 1
not(REIF_259): 0
not(REIF_259): 1
REIF_260: 0
REIF_260: 1
not(REIF_260): 0
not(REIF_260): 1
REIF_261: 0
REIF_261: 1
not(REIF_261): 0
not(REIF_261): 1
REIF_262: 0
REIF_262: 1
not(REIF_262): 0
not(REIF_262): 1
REIF_263: 0
REIF_263: 1
not(REIF_263): 0
not(REIF_263): 1
REIF_264: 0
REIF_264: 1
not(REIF_264): 0
not(REIF_264): 1
REIF_265: 0
REIF_265: 1
not(REIF_265): 0
not(REIF_265): 1
REIF_266: 0
REIF_266: 1
not(REIF_266): 0
not(REIF_266): 1
REIF_267: 0
REIF_267: 1
not(REIF_267): 0
not(REIF_267): 1
REIF_268: 0
REIF_268: 1
not(REIF_268): 0
not(REIF_268): 1
REIF_269: 0
REIF_269: 1
not(REIF_269): 0
not(REIF_269): 1
REIF_270: 0
REIF_270: 1
not(REIF_270): 0
not(REIF_270): 1
REIF_271: 0
REIF_271: 1
not(REIF_271): 0
not(REIF_271): 1
REIF_272: 0
REIF_272: 1
not(REIF_272): 0
not(REIF_272): 1
REIF_273: 0
REIF_273: 1
not(REIF_273): 0
not(REIF_273): 1
REIF_274: 0
REIF_274: 1
not(REIF_274): 0
not(REIF_274): 1
REIF_275: 0
REIF_275: 1
not(REIF_275): 0
not(REIF_275): 1
REIF_276: 0
REIF_276: 1
not(REIF_276): 0
not(REIF_276): 1
REIF_277: 0
REIF_277: 1
not(REIF_277): 0
not(REIF_277): 1
REIF_278: 0
REIF_278: 1
not(REIF_278): 0
not(REIF_278): 1
REIF_279: 0
REIF_279: 1
not(REIF_279): 0
not(REIF_279): 1
REIF_280: 0
REIF_280: 1
not(REIF_280): 0
not(REIF_280): 1
REIF_281: 0
REIF_281: 1
not(REIF_281): 0
not(REIF_281): 1
REIF_282: 0
REIF_282: 1
not(REIF_282): 0
not(REIF_282): 1
REIF_283: 0
REIF_283: 1
not(REIF_283): 0
not(REIF_283): 1
REIF_284: 0
REIF_284: 1
not(REIF_284): 0
not(REIF_284): 1
REIF_285: 0
REIF_285: 1
not(REIF_285): 0
not(REIF_285): 1
REIF_286: 0
REIF_286: 1
not(REIF_286): 0
not(REIF_286): 1
REIF_287: 0
REIF_287: 1
not(REIF_287): 0
not(REIF_287): 1
REIF_288: 0
REIF_288: 1
not(REIF_288): 0
not(REIF_288): 1
REIF_289: 0
REIF_289: 1
not(REIF_289): 0
not(REIF_289): 1
REIF_290: 0
REIF_290: 1
not(REIF_290): 0
not(REIF_290): 1
REIF_291: 0
REIF_291: 1
not(REIF_291): 0
not(REIF_291): 1
REIF_292: 0
REIF_292: 1
not(REIF_292): 0
not(REIF_292): 1
REIF_293: 0
REIF_293: 1
not(REIF_293): 0
not(REIF_293): 1
REIF_294: 0
REIF_294: 1
not(REIF_294): 0
not(REIF_294): 1
REIF_295: 0
REIF_295: 1
not(REIF_295): 0
not(REIF_295): 1
REIF_296: 0
REIF_296: 1
not(REIF_296): 0
not(REIF_296): 1
REIF_297: 0
REIF_297: 1
not(REIF_297): 0
not(REIF_297): 1
REIF_298: 0
REIF_298: 1
not(REIF_298): 0
not(REIF_298): 1
REIF_299: 0
REIF_299: 1
not(REIF_299): 0
not(REIF_299): 1
REIF_300: 0
REIF_300: 1
not(REIF_300): 0
not(REIF_300): 1
REIF_301: 0
REIF_301: 1
not(REIF_301): 0
not(REIF_301): 1
REIF_302: 0
REIF_302: 1
not(REIF_302): 0
not(REIF_302): 1
REIF_303: 0
REIF_303: 1
not(REIF_303): 0
not(REIF_303): 1
REIF_304: 0
REIF_304: 1
not(REIF_304): 0
not(REIF_304): 1
REIF_305: 0
REIF_305: 1
not(REIF_305): 0
not(REIF_305): 1
REIF_306: 0
REIF_306: 1
not(REIF_306): 0
not(REIF_306): 1
REIF_307: 0
REIF_307: 1
not(REIF_307): 0
not(REIF_307): 1
REIF_308: 0
REIF_308: 1
not(REIF_308): 0
not(REIF_308): 1
REIF_309: 0
REIF_309: 1
not(REIF_309): 0
not(REIF_309): 1
REIF_310: 0
REIF_310: 1
not(REIF_310): 0
not(REIF_310): 1
REIF_311: 0
REIF_311: 1
not(REIF_311): 0
not(REIF_311): 1
REIF_312: 0
REIF_312: 1
not(REIF_312): 0
not(REIF_312): 1
REIF_313: 0
REIF_313: 1
not(REIF_313): 0
not(REIF_313): 1
REIF_314: 0
REIF_314: 1
not(REIF_314): 0
not(REIF_314): 1
REIF_315: 0
REIF_315: 1
not(REIF_315): 0
not(REIF_315): 1
REIF_316: 0
REIF_316: 1
not(REIF_316): 0
not(REIF_316): 1
REIF_317: 0
REIF_317: 1
not(REIF_317): 0
not(REIF_317): 1
REIF_318: 0
REIF_318: 1
not(REIF_318): 0
not(REIF_318): 1
REIF_319: 0
REIF_319: 1
not(REIF_319): 0
not(REIF_319): 1
REIF_320: 0
REIF_320: 1
not(REIF_320): 0
not(REIF_320): 1
REIF_321: 0
REIF_321: 1
not(REIF_321): 0
not(REIF_321): 1
REIF_322: 0
REIF_322: 1
not(REIF_322): 0
not(REIF_322): 1
REIF_323: 0
REIF_323: 1
not(REIF_323): 0
not(REIF_323): 1
REIF_324: 0
REIF_324: 1
not(REIF_324): 0
not(REIF_324): 1
REIF_325: 0
REIF_325: 1
not(REIF_325): 0
not(REIF_325): 1
REIF_326: 0
REIF_326: 1
not(REIF_326): 0
not(REIF_326): 1

End

Solution
x[0]=0.0
x[1]=328.0
x[2]=348.0
x[3]=368.0
x[4]=388.0
x[5]=408.0
x[6]=418.0
x[7]=428.0
x[8]=438.0
x[9]=451.0
x[10]=461.0
x[11]=471.0
x[12]=481.0
x[13]=491.0
x[14]=501.0
x[15]=511.0
x[16]=521.0
x[17]=531.0
x[18]=541.0
x[19]=551.0
x[20]=561.0
x[21]=571.0
x[22]=581.0
x[23]=593.0
x[24]=608.0
x[25]=623.0
x[26]=638.0
x[27]=653.0
x[28]=668.0
x[29]=683.0
x[30]=698.0
x[31]=713.0
x[32]=728.0
x[33]=743.0
x[34]=758.0
x[35]=773.0
x[36]=788.0
x[37]=803.0
x[38]=818.0
x[39]=833.0
x[40]=848.0
x[41]=863.0
x[42]=878.0
x[43]=893.0
x[44]=908.0
x[45]=923.0
x[46]=935.0
x[47]=947.0
x[48]=959.0
x[49]=971.0
x[50]=983.0
x[51]=995.0
x[52]=1007.0
x[53]=1019.0
x[54]=1031.0
x[55]=1043.0
x[56]=1055.0
x[57]=1067.0
x[58]=1079.0
x[59]=1091.0
x[60]=1103.0
x[61]=1115.0
x[62]=1127.0
x[63]=1139.0
x[64]=1151.0
x[65]=1163.0
x[66]=1172.0
x[67]=1184.0
x[68]=1197.0
x[69]=1211.0
x[70]=1224.0
x[71]=1238.0
x[72]=1253.0
x[73]=1268.0
x[74]=1283.0
x[75]=1298.0
x[76]=1313.0
x[77]=1328.0
x[78]=1343.0
x[79]=1358.0
x[80]=1378.0
x[81]=1398.0
x[82]=262.0
x[83]=282.0
x[84]=302.0
x[85]=322.0
x[86]=332.0
x[87]=342.0
x[88]=352.0
x[89]=362.0
x[90]=372.0
x[91]=382.0
x[92]=392.0
x[93]=402.0
x[94]=412.0
x[95]=422.0
x[96]=432.0
x[97]=442.0
x[98]=452.0
x[99]=462.0
x[100]=472.0
x[101]=482.0
x[102]=492.0
x[103]=502.0
x[104]=517.0
x[105]=532.0
x[106]=547.0
x[107]=562.0
x[108]=577.0
x[109]=592.0
x[110]=607.0
x[111]=622.0
x[112]=637.0
x[113]=652.0
x[114]=667.0
x[115]=682.0
x[116]=697.0
x[117]=712.0
x[118]=727.0
x[119]=742.0
x[120]=757.0
x[121]=772.0
x[122]=787.0
x[123]=802.0
x[124]=817.0
x[125]=832.0
x[126]=847.0
x[127]=859.0
x[128]=871.0
x[129]=883.0
x[130]=895.0
x[131]=907.0
x[132]=919.0
x[133]=931.0
x[134]=943.0
x[135]=955.0
x[136]=967.0
x[137]=979.0
x[138]=991.0
x[139]=1003.0
x[140]=1015.0
x[141]=1027.0
x[142]=1039.0
x[143]=1051.0
x[144]=1063.0
x[145]=1075.0
x[146]=1087.0
x[147]=1099.0
x[148]=1111.0
x[149]=1125.0
x[150]=1140.0
x[151]=1155.0
x[152]=1170.0
x[153]=1185.0
x[154]=1200.0
x[155]=1215.0
x[156]=1230.0
x[157]=1245.0
x[158]=1260.0
x[159]=1278.0
x[160]=1297.0
x[161]=1317.0
x[162]=1337.0
x[163]=1.0
x[164]=0.0
x[165]=1.0
x[166]=1.0
x[167]=0.0
x[168]=1.0
x[169]=0.0
x[170]=1.0
x[171]=0.0
x[172]=1.0
x[173]=0.0
x[174]=1.0
x[175]=0.0
x[176]=1.0
x[177]=0.0
x[178]=1.0
x[179]=0.0
x[180]=1.0
x[181]=0.0
x[182]=1.0
x[183]=0.0
x[184]=1.0
x[185]=0.0
x[186]=1.0
x[187]=0.0
x[188]=1.0
x[189]=0.0
x[190]=1.0
x[191]=0.0
x[192]=1.0
x[193]=0.0
x[194]=1.0
x[195]=0.0
x[196]=1.0
x[197]=0.0
x[198]=1.0
x[199]=0.0
x[200]=1.0
x[201]=0.0
x[202]=1.0
x[203]=0.0
x[204]=1.0
x[205]=0.0
x[206]=1.0
x[207]=0.0
x[208]=1.0
x[209]=0.0
x[210]=1.0
x[211]=0.0
x[212]=1.0
x[213]=0.0
x[214]=1.0
x[215]=0.0
x[216]=1.0
x[217]=0.0
x[218]=1.0
x[219]=0.0
x[220]=1.0
x[221]=0.0
x[222]=1.0
x[223]=0.0
x[224]=1.0
x[225]=0.0
x[226]=1.0
x[227]=0.0
x[228]=1.0
x[229]=0.0
x[230]=1.0
x[231]=0.0
x[232]=1.0
x[233]=0.0
x[234]=1.0
x[235]=0.0
x[236]=1.0
x[237]=0.0
x[238]=1.0
x[239]=0.0
x[240]=1.0
x[241]=0.0
x[242]=1.0
x[243]=0.0
x[244]=1.0
x[245]=0.0
x[246]=1.0
x[247]=0.0
x[248]=1.0
x[249]=0.0
x[250]=1.0
x[251]=0.0
x[252]=1.0
x[253]=0.0
x[254]=1.0
x[255]=0.0
x[256]=1.0
x[257]=0.0
x[258]=1.0
x[259]=0.0
x[260]=1.0
x[261]=0.0
x[262]=1.0
x[263]=0.0
x[264]=1.0
x[265]=0.0
x[266]=1.0
x[267]=0.0
x[268]=1.0
x[269]=0.0
x[270]=1.0
x[271]=0.0
x[272]=1.0
x[273]=0.0
x[274]=1.0
x[275]=0.0
x[276]=1.0
x[277]=0.0
x[278]=1.0
x[279]=0.0
x[280]=1.0
x[281]=0.0
x[282]=1.0
x[283]=0.0
x[284]=1.0
x[285]=0.0
x[286]=1.0
x[287]=0.0
x[288]=1.0
x[289]=0.0
x[290]=1.0
x[291]=0.0
x[292]=1.0
x[293]=0.0
x[294]=1.0
x[295]=0.0
x[296]=1.0
x[297]=0.0
x[298]=1.0
x[299]=0.0
x[300]=1.0
x[301]=0.0
x[302]=1.0
x[303]=0.0
x[304]=1.0
x[305]=0.0
x[306]=1.0
x[307]=0.0
x[308]=1.0
x[309]=0.0
x[310]=1.0
x[311]=0.0
x[312]=1.0
x[313]=0.0
x[314]=1.0
x[315]=0.0
x[316]=1.0
x[317]=0.0
x[318]=1.0
x[319]=0.0
x[320]=1.0
x[321]=0.0
x[322]=1.0
x[323]=0.0
x[324]=1.0
x[325]=0.0
x[326]=1.0
x[327]=0.0
x[328]=1.0
x[329]=0.0
x[330]=1.0
x[331]=0.0
x[332]=1.0
x[333]=0.0
x[334]=1.0
x[335]=0.0
x[336]=1.0
x[337]=0.0
x[338]=1.0
x[339]=0.0
x[340]=1.0
x[341]=0.0
x[342]=1.0
x[343]=0.0
x[344]=1.0
x[345]=0.0
x[346]=1.0
x[347]=0.0
x[348]=1.0
x[349]=0.0
x[350]=1.0
x[351]=0.0
x[352]=1.0
x[353]=0.0
x[354]=1.0
x[355]=0.0
x[356]=1.0
x[357]=0.0
x[358]=1.0
x[359]=0.0
x[360]=1.0
x[361]=0.0
x[362]=1.0
x[363]=0.0
x[364]=1.0
x[365]=0.0
x[366]=1.0
x[367]=0.0
x[368]=1.0
x[369]=0.0
x[370]=1.0
x[371]=0.0
x[372]=1.0
x[373]=0.0
x[374]=1.0
x[375]=0.0
x[376]=1.0
x[377]=0.0
x[378]=1.0
x[379]=0.0
x[380]=1.0
x[381]=0.0
x[382]=1.0
x[383]=0.0
x[384]=1.0
x[385]=0.0
x[386]=1.0
x[387]=0.0
x[388]=1.0
x[389]=0.0
x[390]=1.0
x[391]=0.0
x[392]=1.0
x[393]=0.0
x[394]=1.0
x[395]=0.0
x[396]=1.0
x[397]=0.0
x[398]=1.0
x[399]=0.0
x[400]=1.0
x[401]=0.0
x[402]=1.0
x[403]=0.0
x[404]=1.0
x[405]=0.0
x[406]=1.0
x[407]=0.0
x[408]=1.0
x[409]=0.0
x[410]=1.0
x[411]=0.0
x[412]=1.0
x[413]=0.0
x[414]=1.0
x[415]=0.0
x[416]=1.0
x[417]=0.0
x[418]=1.0
x[419]=0.0
x[420]=1.0
x[421]=0.0
x[422]=1.0
x[423]=0.0
x[424]=1.0
x[425]=0.0
x[426]=1.0
x[427]=0.0
x[428]=1.0
x[429]=0.0
x[430]=1.0
x[431]=0.0
x[432]=1.0
x[433]=0.0
x[434]=1.0
x[435]=0.0
x[436]=1.0
x[437]=0.0
x[438]=1.0
x[439]=0.0
x[440]=1.0
x[441]=0.0
x[442]=1.0
x[443]=0.0
x[444]=1.0
x[445]=0.0
x[446]=1.0
x[447]=0.0
x[448]=1.0
x[449]=0.0
x[450]=1.0
x[451]=0.0
x[452]=1.0
x[453]=0.0
x[454]=1.0
x[455]=0.0
x[456]=1.0
x[457]=0.0
x[458]=1.0
x[459]=0.0
x[460]=1.0
x[461]=0.0
x[462]=1.0
x[463]=0.0
x[464]=1.0
x[465]=0.0
x[466]=1.0
x[467]=0.0
x[468]=1.0
x[469]=0.0
x[470]=1.0
x[471]=0.0
x[472]=1.0
x[473]=0.0
x[474]=1.0
x[475]=0.0
x[476]=1.0
x[477]=0.0
x[478]=1.0
x[479]=0.0
x[480]=1.0
x[481]=0.0
x[482]=1.0
x[483]=0.0
x[484]=1.0
x[485]=0.0
x[486]=1.0
x[487]=0.0
x[488]=1.0
x[489]=0.0
x[490]=1.0
x[491]=0.0
x[492]=1.0
x[493]=0.0
x[494]=1.0
x[495]=0.0
x[496]=1.0
x[497]=0.0
x[498]=1.0
x[499]=0.0
x[500]=1.0
x[501]=0.0
x[502]=1.0
x[503]=0.0
x[504]=1.0
x[505]=0.0
x[506]=1.0
x[507]=0.0
x[508]=1.0
x[509]=0.0
x[510]=1.0
x[511]=0.0
x[512]=1.0
x[513]=0.0
x[514]=1.0
x[515]=0.0
x[516]=1.0
x[517]=0.0
x[518]=1.0
x[519]=0.0
x[520]=1.0
x[521]=0.0
x[522]=1.0
x[523]=0.0
x[524]=1.0
x[525]=0.0
x[526]=1.0
x[527]=0.0
x[528]=1.0
x[529]=0.0
x[530]=1.0
x[531]=0.0
x[532]=1.0
x[533]=0.0
x[534]=1.0
x[535]=0.0
x[536]=1.0
x[537]=0.0
x[538]=1.0
x[539]=0.0
x[540]=1.0
x[541]=0.0
x[542]=1.0
x[543]=0.0
x[544]=1.0
x[545]=0.0
x[546]=1.0
x[547]=0.0
x[548]=1.0
x[549]=0.0
x[550]=1.0
x[551]=0.0
x[552]=1.0
x[553]=0.0
x[554]=1.0
x[555]=0.0
x[556]=1.0
x[557]=0.0
x[558]=1.0
x[559]=0.0
x[560]=1.0
x[561]=0.0
x[562]=1.0
x[563]=0.0
x[564]=1.0
x[565]=0.0
x[566]=1.0
x[567]=0.0
x[568]=1.0
x[569]=0.0
x[570]=1.0
x[571]=0.0
x[572]=1.0
x[573]=0.0
x[574]=1.0
x[575]=0.0
x[576]=1.0
x[577]=0.0
x[578]=1.0
x[579]=0.0
x[580]=1.0
x[581]=0.0
x[582]=1.0
x[583]=0.0
x[584]=1.0
x[585]=0.0
x[586]=1.0
x[587]=0.0
x[588]=1.0
x[589]=0.0
x[590]=1.0
x[591]=0.0
x[592]=1.0
x[593]=0.0
x[594]=1.0
x[595]=0.0
x[596]=1.0
x[597]=0.0
x[598]=1.0
x[599]=0.0
x[600]=1.0
x[601]=0.0
x[602]=1.0
x[603]=0.0
x[604]=1.0
x[605]=0.0
x[606]=1.0
x[607]=0.0
x[608]=1.0
x[609]=0.0
x[610]=1.0
x[611]=0.0
x[612]=1.0
x[613]=0.0
x[614]=1.0
x[615]=0.0
x[616]=1.0
x[617]=0.0
x[618]=1.0
x[619]=0.0
x[620]=1.0
x[621]=0.0
x[622]=1.0
x[623]=0.0
x[624]=1.0
x[625]=0.0
x[626]=1.0
x[627]=0.0
x[628]=1.0
x[629]=0.0
x[630]=1.0
x[631]=0.0
x[632]=1.0
x[633]=0.0
x[634]=1.0
x[635]=0.0
x[636]=1.0
x[637]=0.0
x[638]=1.0
x[639]=0.0
x[640]=1.0
x[641]=0.0
x[642]=1.0
x[643]=0.0
x[644]=1.0
x[645]=0.0
x[646]=1.0
x[647]=0.0
x[648]=1.0
x[649]=0.0
x[650]=1.0
x[651]=0.0
x[652]=1.0
x[653]=0.0
x[654]=1.0
x[655]=0.0
x[656]=1.0
x[657]=0.0
x[658]=1.0
x[659]=0.0
x[660]=1.0
x[661]=0.0
x[662]=1.0
x[663]=0.0
x[664]=1.0
x[665]=0.0
x[666]=1.0
x[667]=0.0
x[668]=1.0
x[669]=0.0
x[670]=1.0
x[671]=0.0
x[672]=1.0
x[673]=0.0
x[674]=1.0
x[675]=0.0
x[676]=1.0
x[677]=0.0
x[678]=1.0
x[679]=0.0
x[680]=1.0
x[681]=0.0
x[682]=1.0
x[683]=0.0
x[684]=1.0
x[685]=0.0
x[686]=1.0
x[687]=0.0
x[688]=1.0
x[689]=0.0
x[690]=1.0
x[691]=0.0
x[692]=1.0
x[693]=0.0
x[694]=1.0
x[695]=0.0
x[696]=1.0
x[697]=0.0
x[698]=1.0
x[699]=0.0
x[700]=1.0
x[701]=0.0
x[702]=1.0
x[703]=0.0
x[704]=1.0
x[705]=0.0
x[706]=1.0
x[707]=0.0
x[708]=1.0
x[709]=0.0
x[710]=1.0
x[711]=0.0
x[712]=1.0
x[713]=0.0
x[714]=1.0
x[715]=0.0
x[716]=1.0
x[717]=0.0
x[718]=1.0
x[719]=0.0
x[720]=1.0
x[721]=0.0
x[722]=1.0
x[723]=0.0
x[724]=1.0
x[725]=0.0
x[726]=1.0
x[727]=0.0
x[728]=1.0
x[729]=0.0
x[730]=1.0
x[731]=0.0
x[732]=1.0
x[733]=0.0
x[734]=1.0
x[735]=0.0
x[736]=1.0
x[737]=0.0
x[738]=1.0
x[739]=0.0
x[740]=1.0
x[741]=0.0
x[742]=1.0
x[743]=0.0
x[744]=1.0
x[745]=0.0
x[746]=1.0
x[747]=0.0
x[748]=1.0
x[749]=0.0
x[750]=1.0
x[751]=0.0
x[752]=1.0
x[753]=0.0
x[754]=1.0
x[755]=0.0
x[756]=1.0
x[757]=0.0
x[758]=1.0
x[759]=0.0
x[760]=1.0
x[761]=0.0
x[762]=1.0
x[763]=0.0
x[764]=1.0
x[765]=0.0
x[766]=1.0
x[767]=0.0
x[768]=1.0
x[769]=0.0
x[770]=1.0
x[771]=0.0
x[772]=1.0
x[773]=0.0
x[774]=1.0
x[775]=0.0
x[776]=1.0
x[777]=0.0
x[778]=1.0
x[779]=0.0
x[780]=1.0
x[781]=0.0
x[782]=1.0
x[783]=0.0
x[784]=1.0
x[785]=0.0
x[786]=1.0
x[787]=0.0
x[788]=1.0
x[789]=0.0
x[790]=1.0
x[791]=0.0
x[792]=1.0
x[793]=0.0
x[794]=1.0
x[795]=0.0
x[796]=1.0
x[797]=0.0
x[798]=1.0
x[799]=0.0
x[800]=1.0
x[801]=0.0
x[802]=1.0
x[803]=0.0
x[804]=1.0
x[805]=0.0
x[806]=1.0
x[807]=0.0
x[808]=1.0
x[809]=0.0
x[810]=1.0
x[811]=0.0
x[812]=1.0
x[813]=0.0
