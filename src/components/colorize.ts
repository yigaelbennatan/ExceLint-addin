export class Colorize {

    // Matchers for all kinds of Excel expressions.
    private static general_re = '\\$?[A-Z]+\\$?\\d+'; // column and row number, optionally with $
    private static sheet_re = '[^\\!]+';
    private static sheet_plus_cell = new RegExp('('+Colorize.sheet_re+')\\!('+Colorize.general_re+')');
    private static sheet_plus_range = new RegExp('('+Colorize.sheet_re+')\\!('+Colorize.general_re+'):('+Colorize.general_re+')');
    private static single_dep = new RegExp('('+Colorize.general_re+')');
    private static range_pair = new RegExp('('+Colorize.general_re+'):('+Colorize.general_re+')', 'g');
    private static cell_both_relative = new RegExp('[^\\$]?([A-Z]+)(\\d+)');
    private static cell_col_absolute = new RegExp('\\$([A-Z]+)[^\\$\\d]?(\\d+)');
    private static cell_row_absolute = new RegExp('[^\\$]?([A-Z]+)\\$(\\d+)');
    private static cell_both_absolute = new RegExp('\\$([A-Z]+)\\$(\\d+)');

    private static color_list = ["pink", "blue", "seagreen", "green", "darkturquoise", "darkgray", "darksalmon", "mediumvioletred" ];
    private static light_color_list = ["LightPink", "LightBlue", "LightYellow", "LightGreen", "LightSkyBlue", "LightGray", "LightSalmon", "PaleVioletRed" ];
    private static light_color_dict = { "pink" : "LightPink",
					"blue" : "LightBlue",
					"seagreen" : "LightSeaGreen",
					"green" : "PaleGreen",
					"darkturquoise" : "PaleTurquoise",
					"darkgray" : "LightGray",
					"darksalmon" : "LightSalmon",
				        "mediumvioletred" : "PaleVioletRed" };
    
    public static get_color(hashval: number) : string {
	return Colorize.color_list[hashval % Colorize.color_list.length];
    }

    public static get_light_color_version(color: string) : string {
	return Colorize.light_color_dict[color];
	//	return Colorize.light_color_list[hashval % Colorize.color_list.length];
    }

    /*
      private static transpose(array) {
      array[0].map((col, i) => array.map(row => row[i]));
      }
    */
    
    public static process_formulas(formulas: Array<Array<string>>, origin_col : number, origin_row : number) : Array<[[number, number], string]> {
	let output : Array<[[number, number], string]> = [];
	// Build up all of the columns of colors.
	for (let i = 0; i < formulas.length; i++) {
	    let row = formulas[i];
	    for (let j = 0; j < row.length; j++) {
		//	console.log("checking "+row[j]);
		//	console.log("char 0 = " + row[j][0]);
		if ((row[j].length > 0) && (row[j][0] === "=")) {
		    //		    console.log("FOUND ONE formulas["+i+","+j+"] = " + row[j]);
		    let vec = Colorize.dependencies(row[j], j + origin_col, i + origin_row);
		    //console.log(vec);
		    let hash =Colorize.hash_vector(vec);
		    //console.log(hash);
//		    let color = Colorize.get_color(hash);
		    //console.log(color);
		    //		    let dict = { "format" : { "fill" : { "color" : color } } };
		    //		    let cell = Colorize.column_index_to_name(j + origin_col + 1)+(i + origin_row + 1);
		    //		    output.push([i, j, color]);
		    output.push([[j + origin_col + 1, i + origin_row + 1], hash.toString()]);
		}
	    }
	}
	
	return output;
    }

    
    public static color_all_data(formulas: Array<Array<string>>, processed_formulas: Array<[[number, number], string]>, origin_col: number, origin_row: number) {
	let refs = Colorize.generate_all_references(formulas, origin_col, origin_row);
	let data_color = {};
	let processed_data = [];
	
	// Generate all formula colors (as a dict).
	let formula_hash = {};
	for (let f of processed_formulas) {
	    let formula_vec = f[0];
	    formula_hash[formula_vec.join(",")] = f[1];
	}
	
	// Color all references based on the color of their referring formula.
	for (let refvec of Object.keys(refs)) {
	    // console.log("refvec = "+refvec);
	    // console.log("ref loop checking refvec = " + refvec);
	    for (let r of refs[refvec]) {
		// console.log("ref loop checking " + r);
		let hash = formula_hash[r.join(",")];
		if (!(hash === undefined)) {
		    //		    console.log("color = " + color);
		    let rv = JSON.parse("[" + refvec + "]");
		    //console.log(parseInt(rv[0]));
		    //console.log(parseInt(rv[1]));
		    let row = parseInt(rv[0]);
		    let col = parseInt(rv[1]);
		    // console.log("Checking "+row+", "+col);
		    if (!([row,col].join(",") in formula_hash)) {
			if (!([row,col].join(",") in data_color)) {
			    processed_data.push([[row, col], hash]);
			    // currentWorksheet.getCell(col-1, row-1).format.fill.color = Colorize.get_light_color_version(color);
			    data_color[[row,col].join(",")] = hash;
			    // console.log("Added "+row+", "+col);
			}
		    }
		}
	    }
	}
	return processed_data;
    }

    
    private static hash(str: string) : number {
	// From https://github.com/darkskyapp/string-hash
	var hash = 5381,
	i = str.length;
	
	while(i) {
	    hash = (hash * 33) ^ str.charCodeAt(--i);
	}

	/* JavaScript does bitwise operations (like XOR, above) on 32-bit signed
	 * integers. Since we want the results to be always positive, convert the
	 * signed int to an unsigned by doing an unsigned bitshift. */
	return hash >>> 0;
    }
    
    private static rgbFromHSV(h,s,v) : Array<number> {
	// From https://gist.github.com/mjackson/5311256#gistcomment-2789005
	/**
	 * I: An array of three elements hue (h) ∈ [0, 360], and saturation (s) and value (v) which are ∈ [0, 1]
	 * O: An array of red (r), green (g), blue (b), all ∈ [0, 255]
	 * Derived from https://en.wikipedia.org/wiki/HSL_and_HSV
	 * This stackexchange was the clearest derivation I found to reimplement https://cs.stackexchange.com/questions/64549/convert-hsv-to-rgb-colors
	 */

	let hprime = h / 60;
	const c = v * s;
	const x = c * (1 - Math.abs(hprime % 2 - 1)); 
	const m = v - c;
	let r, g, b;
	if (!hprime) {r = 0; g = 0; b = 0; }
	if (hprime >= 0 && hprime < 1) { r = c; g = x; b = 0}
	if (hprime >= 1 && hprime < 2) { r = x; g = c; b = 0}
	if (hprime >= 2 && hprime < 3) { r = 0; g = c; b = x}
	if (hprime >= 3 && hprime < 4) { r = 0; g = x; b = c}
	if (hprime >= 4 && hprime < 5) { r = x; g = 0; b = c}
	if (hprime >= 5 && hprime < 6) { r = c; g = 0; b = x}
	
	r = Math.round( (r + m)* 255);
	g = Math.round( (g + m)* 255);
	b = Math.round( (b + m)* 255);

	return [r, g, b]
    }
    
    
    // Convert an Excel column name (a string of alphabetical charcaters) into a number.
    public static column_name_to_index(name: string) : number {
	if (name.length === 1) { // optimizing for the overwhelmingly common case
	    return name[0].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
	}
	let value = 0;
	let reversed_name = name.split("").reverse();
	for (let i of reversed_name) {
	    value *= 26;
	    value = (i.charCodeAt(0) - 'A'.charCodeAt(0)) + 1;
	}
	return value;
    }

    // Convert a column number to a name (as in, 3 => "C").
    public static column_index_to_name(index: number) : string {
	let str = "";
	while (index > 0) {
	    str += String.fromCharCode((index - 1) % 26 + 65); // 65 = 'A'
	    index = Math.floor(index / 26);
	}
	return str.split("").reverse().join("");
    }

    // Take in a list of [[row, col], color] pairs and group them,
    // sorting them (e.g., by columns).
    private static identify_ranges(list : Array<[[number, number], string]>,
				   sortfn? : (n1: [number, number], n2: [number, number]) => number )
    : { [val : string] : Array<[number, number]> }
    {
	// Separate into groups based on their string value.
	let groups = {};
	for (let r of list) {
	    groups[r[1]] = groups[r[1]] || [];
	    groups[r[1]].push(r[0]);
	}
	// Now sort them all.
	for (let k of Object.keys(groups)) {
	    //	console.log(k);
	    groups[k].sort(sortfn);
	    //	console.log(groups[k]);
	}
	return groups;
    }

    private static group_ranges(groups : { [val : string] : Array<[number, number]> },
				columnFirst: boolean)
    : { [val : string] : Array<[[number, number], [number, number]]> }
    {
	let output = {};
	let index0 = 0; // column
	let index1 = 1; // row
	if (!columnFirst) {
	    index0 = 1; // row
	    index1 = 0; // column
	}
	for (let k of Object.keys(groups)) {
	    output[k] = [];
	    let prev = groups[k].shift();
	    let last = prev;
	    for (let v of groups[k]) {
		// Check if in the same column, adjacent row (if columnFirst; otherwise, vice versa).
		if ((v[index0] === last[index0]) && (v[index1] === last[index1] + 1)) {
		    last = v;
		} else {
		    output[k].push([prev, last]);
		    prev = v;
		    last = v;
		}
	    }
	    output[k].push([prev, last]);
	}

	/*
	  let output2 = {};

	  // Need to sort here by row... FIXME
	  
	  for (let k of Object.keys(output)) {
	  output2[k] = [];
	  let prev = output[k].shift();
	  let last = prev;
	  for (let v of output[k]) {
	  if ((v[0] === last[0] + 1) && (v[1] === last[1])) { // same row, adjacent column
	  last = v;
	  } else {
	  output2[k].push([prev, last]);
	  prev = v;
	  last = v;
	  }
	  }
	  output2[k].push([prev, last]);
	  }
	  return output2;*/
	return output;
    }

    public static identify_groups(list : Array<[[number, number], string]>) : { [val : string] : Array<[[number, number], [number, number]]> }
    {
	console.log("start identify_groups");
	console.log(list);
	let columnsort = (a, b) => { if (a[0] == b[0]) { return a[1] - b[1]; } else { return a[0] - b[0]; }};
	let id = Colorize.identify_ranges(list, columnsort);
	let gr = Colorize.group_ranges(id, true); // column-first
	console.log("group ranges");
	console.log(gr);
	// Now try to merge stuff with the same hash.
	let newGr1 = JSON.parse(JSON.stringify(gr)); // deep copy
	let newGr2 = JSON.parse(JSON.stringify(gr)); // deep copy
	let mr = Colorize.mergeable(newGr1);
	console.log("mergeable!");
	console.log(mr);
	let mg = Colorize.merge_groups(newGr2, mr);
	console.log("merge_groups!");
	console.log(mg);
	console.log("end identify_groups");
	return gr;
    }
    

    // True if combining A and B would result in a new rectangle.
    public static merge_friendly(A : [[number,number], [number,number]], B: [[number,number], [number,number]]) : boolean {
	let [[Ax0, Ay0], [Ax1, Ay1]] = A;
	let [[Bx0, By0], [Bx1, By1]] = B;
	if ((Ax0 == Bx0) && (Ax1 == Bx1)) {
	    if (Ay0 == By1 + 1) {
		// top
		return true;
	    }
	    if (Ay1 + 1 == By0) {
		// bottom
		return true;
	    }
	}
	if ((Ay0 == By0) && (Ay1 == By1)) {
	    if (Ax0 == Bx1 + 1) {
		// left
		return true;
	    }
	    if (Ax1 + 1 == Bx0) {
		// right
		return true;
	    }
	}
	return false;
    }

    // Return a merged version (both should be "merge friendly").
    public static merge_rectangles(A : [[number,number], [number,number]],
				   B: [[number,number], [number,number]])
    : [[number, number], [number, number]]
    {
	let [[Ax0, Ay0], [Ax1, Ay1]] = A;
	let [[Bx0, By0], [Bx1, By1]] = B;
	if ((Ax0 == Bx0) && (Ax1 == Bx1)) {
	    if (Ay0 == By1 + 1) {
		// top
		return [[Bx0, By0], [Ax0, Ay1]];
	    }
	    if (Ay1 + 1 == By0) {
		// bottom
		return [[Ax0, Ay0], [Bx1, By1]];
	    }
	}
	if ((Ay0 == By0) && (Ay1 == By1)) {
	    if (Ax0 == Bx1 + 1) {
		// left
		return [[Bx0, By0], [Ax1, Ay1]];
	    }
	    if (Ax1 + 1 == Bx0) {
		// right
		return [[Ax0, Ay0], [Bx1, By1]];
	    }
	}
	return [[-1, -1], [-1, -1]]; //FIXME should throw an exception here
    }

    public static merge_groups(groups : { [val : string] : Array<[[number, number], [number, number]]> },
			       merge_candidates : { [val: string] : Array<Array<[[number, number], [number, number]]>> })
    : { [val : string] : Array<[[number, number], [number, number]]> }
    {
	// Groups already passed as input to mergeable.
	// Merge_candidates generated by mergeable.
	// Go through all mergeable groups; for each, remove the corresponding two rectangles and add the merged one.
	let merged_rectangles = {}
	for (let k of Object.keys(merge_candidates)) {
	    merged_rectangles[k] = merged_rectangles[k] || [];
	    let removed = {};
	    for (let range of merge_candidates[k]) {
		let first = range[0];
		let second = range[1];
		// Add these to be removed later.
		removed[JSON.stringify(first)] = true;
		removed[JSON.stringify(second)] = true;
		let merged = Colorize.merge_rectangles(first, second);
		merged_rectangles[k].push(merged);
	    }
	    let newList = [];
	    for (let i = 0; i < groups[k].length; i++) {
		let str = JSON.stringify(groups[k][i]);
		if (!(str in removed)) {
		    newList.push(groups[k][i]);
		}
	    }
	    merged_rectangles[k].push(...newList);
	}
	return merged_rectangles;
    }
    
    public static mergeable(grouped_ranges: { [val : string] : Array<[[number, number], [number, number]]> })
    : { [val: string] : Array<Array<[[number, number], [number, number]]>> }  {
	// Input comes from group_ranges.
	let mergeable = {};
	for (let k of Object.keys(grouped_ranges)) {
	    mergeable[k] = [];
	    let r = grouped_ranges[k];
	    while (r.length > 0) {
		let head = r.shift();
		let merge_candidates = r.filter((b) => { return Colorize.merge_friendly(head, b); });
		if (merge_candidates.length > 0) {
		    for (let c of merge_candidates) {
			mergeable[k].push([head, c]);
		    }
		}
	    }
	}
	return mergeable;
    }
    
    // Returns a vector (x, y) corresponding to the column and row of the computed dependency.
    public static cell_dependency(cell: string, origin_col: number, origin_row: number) : [number, number] {
	{
	    let r = Colorize.cell_col_absolute.exec(cell);
	    if (r) {
		//	    console.log(JSON.stringify(r));
		let col = Colorize.column_name_to_index(r[1]);
		let row = parseInt(r[2]);
		//	    console.log("absolute col: " + col + ", row: " + row);
		return [col, row - origin_row];
	    }
	}

	{
	    let r = Colorize.cell_both_relative.exec(cell);
	    if (r) {
		//	    console.log("both_relative");
		let col = Colorize.column_name_to_index(r[1]);
		let row = parseInt(r[2]);
		return [col - origin_col, row - origin_row];
	    }
	}

	{
	    let r = Colorize.cell_row_absolute.exec(cell);
	    if (r) {
		//	    console.log("row_absolute");
		let col = Colorize.column_name_to_index(r[1]);
		let row = parseInt(r[2]);
		return [col - origin_col, row];
	    }
	}

	{
	    let r = Colorize.cell_both_absolute.exec(cell);
	    if (r) {
		//	    console.log("both_absolute");
		let col = Colorize.column_name_to_index(r[1]);
		let row = parseInt(r[2]);
		return [col, row];
	    }
	}
	
	throw new Error('We should never get here.');
	return [0, 0];
    }


    public static all_cell_dependencies(range: string) /* , origin_col: number, origin_row: number) */ : Array<[number, number]> {
	
	let found_pair = null;
	let all_vectors : Array<[number, number]> = [];
	
	/// FIX ME - should we count the same range multiple times? Or just once?
	
	// First, get all the range pairs out.
	while (found_pair = Colorize.range_pair.exec(range)) {
	    if (found_pair) {
		//		console.log("all_cell_dependencies --> " + found_pair);
		let first_cell = found_pair[1];
		//		console.log(" first_cell = " + first_cell);
		let first_vec = Colorize.cell_dependency(first_cell, 0, 0);
		//		console.log(" first_vec = " + JSON.stringify(first_vec));
		let last_cell = found_pair[2];
		//		console.log(" last_cell = " + last_cell);
		let last_vec = Colorize.cell_dependency(last_cell, 0, 0);
		//		console.log(" last_vec = " + JSON.stringify(last_vec));

		// First_vec is the upper-left hand side of a rectangle.
		// Last_vec is the lower-right hand side of a rectangle.

		// Generate all vectors.
		let length = last_vec[0] - first_vec[0] + 1;
		let width = last_vec[1] - first_vec[1] + 1;
		for (let x = 0; x < length; x++) {
		    for (let y = 0; y < width; y++) {
			// console.log(" pushing " + (x + first_vec[0]) + ", " + (y + first_vec[1]));
			// console.log(" (x = " + x + ", y = " + y);
			all_vectors.push([x + first_vec[0], y + first_vec[1]]);
		    }
		}
		
		// Wipe out the matched contents of range.
		let newRange = range.replace(found_pair[0], '_'.repeat(found_pair[0].length));
		range = newRange;
	    }
	}

	// Now look for singletons.
	let singleton = null;
	while (singleton = Colorize.single_dep.exec(range)) {
	    if (singleton) {
		//		console.log("SINGLETON");
		//		console.log("singleton[1] = " + singleton[1]);
		//	    console.log(found_pair);
		let first_cell = singleton[1];
		//		console.log(first_cell);
		let vec = Colorize.cell_dependency(first_cell, 0, 0);
		all_vectors.push(vec);
		// Wipe out the matched contents of range.
		let newRange = range.replace(singleton[0], '_'.repeat(singleton[0].length));
		range = newRange;
	    }
	}

	return all_vectors;

    }
    
    public static dependencies(range: string, origin_col: number, origin_row: number) : Array<number> {

	let base_vector = [0, 0];
	
	let found_pair = null;

	/// FIX ME - should we count the same range multiple times? Or just once?
	
	// First, get all the range pairs out.
	while (found_pair = Colorize.range_pair.exec(range)) {
	    if (found_pair) {
		//	    console.log(found_pair);
		let first_cell = found_pair[1];
		//		console.log(first_cell);
		let first_vec = Colorize.cell_dependency(first_cell, origin_col, origin_row);
		let last_cell = found_pair[2];
		//		console.log(last_cell);
		let last_vec = Colorize.cell_dependency(last_cell, origin_col, origin_row);

		// First_vec is the upper-left hand side of a rectangle.
		// Last_vec is the lower-right hand side of a rectangle.
		// Compute the appropriate vectors to be added.

		// e.g., [3, 2] --> [5, 5] ==
		//          [3, 2], [3, 3], [3, 4], [3, 5]
		//          [4, 2], [4, 3], [4, 4], [4, 5]
		//          [5, 2], [5, 3], [5, 4], [5, 5]
		// 
		// vector to be added is [4 * (3 + 4 + 5), 3 * (2 + 3 + 4 + 5) ]
		//  = [48, 42]

		let sum_x = 0;
		let sum_y = 0;
		let width = last_vec[1] - first_vec[1] + 1;   // 4
		sum_x = width * ((last_vec[0]*(last_vec[0]+1))/2 - ((first_vec[0]-1)*((first_vec[0]-1)+1))/2);
		let length = last_vec[0] - first_vec[0] + 1;   // 3
		sum_y = length * ((last_vec[1]*(last_vec[1]+1))/2 - ((first_vec[1]-1)*((first_vec[1]-1)+1))/2);

		base_vector[0] += sum_x;
		base_vector[1] += sum_y;
		
		// Wipe out the matched contents of range.
		let newRange = range.replace(found_pair[0], '_'.repeat(found_pair[0].length));
		range = newRange;
	    }
	}

	// Now look for singletons.
	let singleton = null;
	while (singleton = Colorize.single_dep.exec(range)) {
	    if (singleton) {
		//	    console.log(found_pair);
		let first_cell = singleton[1];
		//		console.log(first_cell);
		let vec = Colorize.cell_dependency(first_cell, origin_col, origin_row);
		base_vector[0] += vec[0];
		base_vector[1] += vec[1];
		// Wipe out the matched contents of range.
		let newRange = range.replace(singleton[0], '_'.repeat(singleton[0].length));
		range = newRange;
	    }
	}

	return base_vector;

    }

    public static generate_all_references(formulas: Array<Array<string>>, origin_col : number, origin_row : number) : { [ dep: string ] : Array<[number, number]> } {
	// Generate all references.
	let refs = {};
	for (let i = 0; i < formulas.length; i++) {
	    let row = formulas[i];
	    for (let j = 0; j < row.length; j++) {
		// console.log("origin_col = "+origin_col+", origin_row = " + origin_row);
		let all_deps = Colorize.all_cell_dependencies(row[j]); // , origin_col + j, origin_row + i);
		if (all_deps.length > 0) {
		    // console.log(all_deps);
		    let src = [origin_col+j, origin_row+i];
		    // console.log("src = " + src);
		    for (let dep of all_deps) {
			let dep2 = dep; // [dep[0]+origin_col, dep[1]+origin_row];
			//				console.log("dep type = " + typeof(dep));
			//				console.log("dep = "+dep);
			refs[dep2.join(",")] = refs[dep2.join(",")] || [];
			refs[dep2.join(",")].push(src);
			// console.log("refs[" + dep2.join(",") + "] = " + JSON.stringify(refs[dep2.join(",")]));
		    }
		}
	    }
	}
	return refs;
    }
    
    
    public static extract_sheet_cell(str: string) : Array<string> {
	let matched = Colorize.sheet_plus_cell.exec(str);
	if (matched) {
	    return [matched[1], matched[2], matched[3]];
	}
	return ["", "", ""];
    }
    
    public static extract_sheet_range(str: string) : Array<string> {
	let matched = Colorize.sheet_plus_range.exec(str);
	if (matched) {
	    return [matched[1], matched[2], matched[3]];
	}
	return ["", "", ""];
    }

    
    public static hash_vector(vec: Array<number>) : number {
	// Return a hash of the given vector.
	let h = Colorize.hash(JSON.stringify(vec));
	return h;
    }
    

}

//console.log(Colorize.dependencies('$C$2:$E$5', 10, 10));
//console.log(Colorize.dependencies('$A$123,A1:B$12,$A12:$B$14', 10, 10));
//console.log(Colorize.hash_vector(Colorize.dependencies('$C$2:$E$5', 10, 10)));
