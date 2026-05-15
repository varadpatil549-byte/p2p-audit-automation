let
  // ================= 1) READ TABLES =================
CelonisRaw = Excel.CurrentWorkbook(){[Name="tblCelonis"]}[Content],
SAPRaw = Excel.CurrentWorkbook(){[Name="tblSAP"]}[Content],
ThresholdsRaw = Excel.CurrentWorkbook(){[Name="tblThresholds"]}[Content],
Vendor100Raw = Excel.CurrentWorkbook(){[Name="tblVendor100"]}[Content],
 
  // Keep original Celonis header order (we'll restore it later)
  CelonisOrigCols = Table.ColumnNames(CelonisRaw),
 
  // ================= 2) STANDARDIZE HEADERS =================
  ToLowerHeaders = (t as table) as table =>
      Table.TransformColumnNames(t, each Text.Lower(Text.Trim(_))),
  Celonis = ToLowerHeaders(CelonisRaw),
  SAP     = ToLowerHeaders(SAPRaw),
  Th      = ToLowerHeaders(ThresholdsRaw),
 
  // Ensure Company Code is text across all 3 tables
  AsText = (col as text) => each if Value.Is(_, type text) then _ else Text.From(_),
  Celonis1 = if Table.HasColumns(Celonis, "company code") then Table.TransformColumns(Celonis, {{"company code", AsText("company code")}}) else error "tblCelonis must contain 'Company Code'.",
  SAP1     = if Table.HasColumns(SAP,     "company code") then Table.TransformColumns(SAP,     {{"company code", AsText("company code")}}) else error "tblSAP must contain 'Company Code'.",
  Th1      = if Table.HasColumns(Th,      "company code") then Table.TransformColumns(Th,      {{"company code", AsText("company code")}}) else error "tblThresholds must contain 'Company Code'.",
 
  // ================= 3) NORMALIZE DOCUMENT IDS =================
  // e.g. "000005325893" -> "5325893"; non-digits removed; blank => null
  NormalizeId = (x as any) as nullable text =>
      let s  = if x=null then null else Text.Select(Text.From(x), {"0".."9"}),
          nz = if s=null then null else Text.TrimStart(s, "0"),
          out = if nz=null or nz="" then null else nz
      in out,
 
  EnsureDocId = (t as table, src as text) as table =>
      if Table.HasColumns(t, "document id") then Table.AddColumn(t, "doc_key", each NormalizeId([document id]), type text)
      else error (src & " must contain 'Document id'."),
 
  Celonis2 = EnsureDocId(Celonis1, "tblCelonis"),
  SAP2     = EnsureDocId(SAP1,     "tblSAP"),
 
  // ================= 4) PICK SAP COLUMNS (ROBUST) =================
  colsSAP      = Table.ColumnNames(SAP2),
  colsSAPLower = List.RemoveNulls(List.Transform(colsSAP, each try Text.Lower(Text.From(_)) otherwise null)),
 
  PickCol = (cands as list, purpose as text) as text =>
      let
        candsLower = List.RemoveNulls(List.Transform(cands, each try Text.Lower(Text.From(_)) otherwise null)),
        hits       = List.Intersect({candsLower, colsSAPLower}),
        result     = if List.Count(hits) > 0 then List.First(hits) else null
      in
        if result <> null then result
        else error ("Couldn't find the " & purpose & " column in tblSAP. Tried: "
                    & Text.Combine(cands, ", ") & ". Available: " & Text.Combine(colsSAP, ", ")),
 
  // Add any local header variants here if needed
  SAP_VendorCol = PickCol({"vendor name","vendor","name"},                                     "Vendor Name"),
  SAP_RefCol    = PickCol({"reference","ref","reference number"},                              "Reference"),
  SAP_DueCol    = PickCol({"due date","baseline date","net due date"},                         "Due Date"),
  // your case: "total amount in rep. crcy" included
  SAP_AmtCol    = PickCol({"total amount in rep. crcy","gross invoice amount","net amount","amount","doc amount","amount in doc. curr.","amount in doc curr"}, "Amount"),
 
  // ================= 5) INNER JOIN (only docs present in SAP survive) =================
  MergeSAP = Table.NestedJoin(
      Celonis2, {"doc_key","company code"},
      SAP2,     {"doc_key","company code"},
      "SAP", JoinKind.Inner),
 
  // Expand with SAFE names (suffix) to avoid collisions with Celonis
  ExpandSAP = Table.ExpandTableColumn(
      MergeSAP, "SAP",
      {SAP_AmtCol, SAP_VendorCol, SAP_RefCol, SAP_DueCol},
      {"amount__sap","vendor name__sap","reference__sap","due date__sap"}),
 
  // ================= 6) TYPES =================
  ToNumber = (x as any) as nullable number =>
      try if Value.Is(x, type number) then Number.From(x)
          else if x=null then null
          else Number.From(Text.Select(Text.From(x), {"0".."9",".","-"}))
      otherwise null,
 
  ToDate = (x as any) as nullable date =>
      try if Value.Is(x, type date) then x
          else if Value.Is(x, type datetime) then Date.From(x)
          else if x=null then null
          else Date.From(DateTime.FromText(Text.From(x)))
      otherwise null,
 
  Typed = Table.TransformColumns(
      ExpandSAP,
      {{"amount__sap", ToNumber, type number}, {"due date__sap", ToDate, type date}}),
 
  // ================= 7) OVERWRITE/ADD CELONIS COLUMNS USING SAP DATA =================
  // Replace Celonis column with SAP column (or add if Celonis didn't have it)
  ReplaceWithSAP = (tbl as table, celName as text, sapName as text) as table =>
      let
        hasSAP  = Table.HasColumns(tbl, sapName),
        step1   = if hasSAP and Table.HasColumns(tbl, celName) then Table.RemoveColumns(tbl, {celName}) else tbl,
        step2   = if hasSAP then Table.RenameColumns(step1, {{sapName, celName}}, MissingField.Ignore) else step1
      in
        step2,
 
  AfterVendor = ReplaceWithSAP(Typed, "vendor name", "vendor name__sap"),
  AfterRef    = ReplaceWithSAP(AfterVendor, "reference",  "reference__sap"),
  AfterDue    = ReplaceWithSAP(AfterRef,    "due date",   "due date__sap"),
  AfterAmt    = ReplaceWithSAP(AfterDue,    "amount",     "amount__sap"),
 
  // ================= 8) THRESHOLDS (incl. "100%") =================
  AddIs100 = Table.AddColumn(Th1, "is_100",
              each Text.EndsWith(Text.Upper(Text.Trim(Text.From([threshold]))), "%"), type logical),
  OnlyDigits = (s as nullable text) as nullable text =>
              if s=null then null else Text.Select(s, {"0".."9",".","-"}),
  AddThrNum = Table.AddColumn(AddIs100, "threshold_num",
              each if [is_100] then null else try Number.From(OnlyDigits(Text.From([threshold]))) otherwise null, type number),

                // ================= 8B) SPECIAL VENDOR 100% AUDIT =================
Vendor100 = Table.TransformColumnNames(Vendor100Raw, each Text.Lower(Text.Trim(_))),
VendorList = List.Transform(Vendor100[vendor], each Text.From(_)),
 
  // ================= 9) MERGE THRESHOLDS (NO NAME CLASH) =================
  MergeTh = Table.NestedJoin(AfterAmt, {"company code"}, AddThrNum, {"company code"}, "TH", JoinKind.LeftOuter),
 
  // Expand with *_thr suffix so we never clash with Celonis' rac/sub region
  ExpandTh = Table.ExpandTableColumn(
      MergeTh, "TH",
      {"threshold","is_100","threshold_num","rac region","sub region"},
      {"threshold","is_100","threshold_num","rac region__thr","sub region__thr"}),
 
  // If Celonis already had RAC/Sub Region, drop *_thr; else rename *_thr to standard names
  KeepRegion =
    let
      hasCelRac = Table.HasColumns(ExpandTh, "rac region"),
      hasCelSub = Table.HasColumns(ExpandTh, "sub region"),
      t1 = if hasCelRac then Table.RemoveColumns(ExpandTh, {"rac region__thr"})
           else Table.RenameColumns(ExpandTh, {{"rac region__thr","rac region"}}, MissingField.Ignore),
      t2 = if hasCelSub then Table.RemoveColumns(t1, {"sub region__thr"})
           else Table.RenameColumns(t1, {{"sub region__thr","sub region"}}, MissingField.Ignore)
    in
      t2,
 
  // ================= 10) AUDIT SCOPE =================
  AddAuditFlag = Table.AddColumn(KeepRegion, "Audit_In_Scope", each
     if [vendor] <> null and List.Contains(VendorList, Text.From([vendor])) then true
     else if [is_100] then true
     else if [threshold_num] <> null and [amount] <> null and [amount] >= [threshold_num] then true
     else false, type logical),
 
  AddReason = Table.AddColumn(AddAuditFlag, "Reason", each
     if [vendor] <> null and List.Contains(VendorList, Text.From([vendor])) then "Vendor 100% audit"
     else if [is_100] then "100% company code"
     else if [threshold_num] = null then "No threshold configured"
     else if [amount] = null then "No amount in SAP"
     else if [amount] >= [threshold_num] then "Amount ≥ Threshold"
     else "Below threshold", type text),
 
  // ================= 11) FINAL ORDER: Celonis columns first =================
  lowerOrig   = List.Transform(CelonisOrigCols, each Text.Lower(_)),
  currentCols = Table.ColumnNames(AddReason),
 
  // Put Celonis columns (original order) first…
  baseOrder   = List.Intersect({lowerOrig, currentCols}),
  // …then everything else that wasn’t in Celonis (extras)
  extras      = List.RemoveItems(currentCols, baseOrder),
  finalOrder  = baseOrder & extras,
 
  Reorder = Table.ReorderColumns(AddReason, finalOrder, MissingField.UseNull),
    #"Changed Type" = Table.TransformColumnTypes(Reorder,{{"date", type date}, {"due date", type date}})
in
  #"Changed Type"