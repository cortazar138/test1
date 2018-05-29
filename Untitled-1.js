function TxtCheck(id, txt) {
	var title = session.findById(id).text;
	return title.search(txt) != -1;
}

function findRow(col, txt, wnd=1, start=0) {
	var id = "wnd[" + wnd + "]/usr/cntlGRID/shellcont/shell"
	var max = session.findById("wnd[" + wnd + "]/usr/cntlGRID/shellcont/shell").maxRows;
	for (i=start; i < max-1; i++) {
		if (session.findById(id).getCellValue(i, col) == txt) {
			return i
			break;
		}
	}
return false;	
}

function selectRow(col, txt, wnd=1, start=0) {
	debugger;
	var id = "wnd[" + wnd + "]/usr/cntlGRID/shellcont/shell"
	var max = session.findById("wnd[" + wnd + "]/usr/cntlGRID/shellcont/shell").maxRows;
	for (i=start; i < max-1; i++) {
		if (session.findById(id).getCellValue(i, col) == txt) {
			session.findById(id).doubleClick(i, col);
			break;
		}
	}
return false;	
}