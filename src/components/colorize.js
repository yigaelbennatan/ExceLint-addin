"use strict";
exports.__esModule = true;
var colorutils_1 = require("./colorutils");
var excelutils_1 = require("./excelutils");
var rectangleutils_1 = require("./rectangleutils");
var Colorize = /** @class */ (function () {
    function Colorize() {
    }
    Colorize.initialize = function () {
        if (!Colorize.initialized) {
            Colorize.make_light_color_versions();
            for (var _i = 0, _a = Object.keys(Colorize.light_color_dict); _i < _a.length; _i++) {
                var i = _a[_i];
                Colorize.color_list.push(i);
                Colorize.light_color_list.push(Colorize.light_color_dict[i]);
            }
            Colorize.initialized = true;
        }
    };
    Colorize.get_color = function (hashval) {
        return Colorize.color_list[hashval % Colorize.color_list.length];
    };
    Colorize.make_light_color_versions = function () {
        //		console.log('YO');
        for (var i = 0; i < 255; i += 7) {
            var rgb = colorutils_1.ColorUtils.HSVtoRGB(i / 255.0, .5, .85);
            var _a = rgb.map(function (x) { return Math.round(x).toString(16).padStart(2, '0'); }), rs = _a[0], gs = _a[1], bs = _a[2];
            var str = '#' + rs + gs + bs;
            str = str.toUpperCase();
            Colorize.light_color_dict[str] = '';
        }
        for (var color in Colorize.light_color_dict) {
            var lightstr = colorutils_1.ColorUtils.adjust_brightness(color, 4.0);
            var darkstr = color; // = Colorize.adjust_color(color, 0.25);
            //			console.log(str);
            //			console.log('Old RGB = ' + color + ', new = ' + str);
            delete Colorize.light_color_dict[color];
            Colorize.light_color_dict[darkstr] = lightstr;
        }
    };
    Colorize.get_light_color_version = function (color) {
        return Colorize.light_color_dict[color];
    };
    /*
      private static transpose(array) {
      array[0].map((col, i) => array.map(row => row[i]));
      }
    */
    Colorize.process_formulas = function (formulas, origin_col, origin_row) {
        var output = [];
        // Build up all of the columns of colors.
        for (var i = 0; i < formulas.length; i++) {
            var row = formulas[i];
            for (var j = 0; j < row.length; j++) {
                if ((row[j].length > 0) && (row[j][0] === '=')) {
                    var vec = excelutils_1.ExcelUtils.dependencies(row[j], j + origin_col, i + origin_row);
                    var hash = Colorize.hash_vector(vec);
                    output.push([[j + origin_col + 1, i + origin_row + 1], hash.toString()]);
                }
            }
        }
        return output;
    };
    Colorize.color_all_data = function (formulas, processed_formulas, origin_col, origin_row) {
        //console.log('color_all_data');
        var refs = Colorize.generate_all_references(formulas, origin_col, origin_row);
        var data_color = {};
        var processed_data = [];
        // Generate all formula colors (as a dict).
        var formula_hash = {};
        for (var _i = 0, processed_formulas_1 = processed_formulas; _i < processed_formulas_1.length; _i++) {
            var f = processed_formulas_1[_i];
            var formula_vec = f[0];
            formula_hash[formula_vec.join(',')] = f[1];
        }
        // Color all references based on the color of their referring formula.
        for (var _a = 0, _b = Object.keys(refs); _a < _b.length; _a++) {
            var refvec = _b[_a];
            for (var _c = 0, _d = refs[refvec]; _c < _d.length; _c++) {
                var r_1 = _d[_c];
                var hash = formula_hash[r_1.join(',')];
                if (!(hash === undefined)) {
                    var rv = JSON.parse('[' + refvec + ']');
                    var row = parseInt(rv[0], 10);
                    var col = parseInt(rv[1], 10);
                    var rj = [row, col].join(',');
                    if (!(rj in formula_hash)) {
                        if (!(rj in data_color)) {
                            processed_data.push([[row, col], hash]);
                            data_color[rj] = hash;
                        }
                    }
                }
            }
        }
        return processed_data;
    };
    Colorize.hash = function (str) {
        // From https://github.com/darkskyapp/string-hash
        var hash = 5381, i = str.length;
        while (i) {
            hash = (hash * 33) ^ str.charCodeAt(--i);
        }
        /* JavaScript does bitwise operations (like XOR, above) on 32-bit signed
        * integers. Since we want the results to be always positive, convert the
        * signed int to an unsigned by doing an unsigned bitshift. */
        return hash >>> 0;
    };
    // Take in a list of [[row, col], color] pairs and group them,
    // sorting them (e.g., by columns).
    Colorize.identify_ranges = function (list, sortfn) {
        // Separate into groups based on their string value.
        var groups = {};
        for (var _i = 0, list_1 = list; _i < list_1.length; _i++) {
            var r_2 = list_1[_i];
            groups[r_2[1]] = groups[r_2[1]] || [];
            groups[r_2[1]].push(r_2[0]);
        }
        // Now sort them all.
        for (var _a = 0, _b = Object.keys(groups); _a < _b.length; _a++) {
            var k = _b[_a];
            //	console.log(k);
            groups[k].sort(sortfn);
            //	console.log(groups[k]);
        }
        return groups;
    };
    Colorize.group_ranges = function (groups, columnFirst) {
        var output = {};
        var index0 = 0; // column
        var index1 = 1; // row
        if (!columnFirst) {
            index0 = 1; // row
            index1 = 0; // column
        }
        for (var _i = 0, _a = Object.keys(groups); _i < _a.length; _i++) {
            var k = _a[_i];
            output[k] = [];
            var prev = groups[k].shift();
            var last = prev;
            for (var _b = 0, _c = groups[k]; _b < _c.length; _b++) {
                var v = _c[_b];
                // Check if in the same column, adjacent row (if columnFirst; otherwise, vice versa).
                if ((v[index0] === last[index0]) && (v[index1] === last[index1] + 1)) {
                    last = v;
                }
                else {
                    output[k].push([prev, last]);
                    prev = v;
                    last = v;
                }
            }
            output[k].push([prev, last]);
        }
        return output;
    };
    Colorize.identify_groups = function (list) {
        var columnsort = function (a, b) { if (a[0] === b[0]) {
            return a[1] - b[1];
        }
        else {
            return a[0] - b[0];
        } };
        var id = Colorize.identify_ranges(list, columnsort);
        var gr = Colorize.group_ranges(id, true); // column-first
        // Now try to merge stuff with the same hash.
        var newGr1 = JSON.parse(JSON.stringify(gr)); // deep copy
        //        let newGr2 = JSON.parse(JSON.stringify(gr)); // deep copy
        //        console.log('group');
        //        console.log(JSON.stringify(newGr1));
        var mg = Colorize.merge_groups(newGr1);
        //        let mr = Colorize.mergeable(newGr1);
        //        console.log('mergeable');
        //       console.log(JSON.stringify(mr));
        //       let mg = Colorize.merge_groups(newGr2, mr);
        //        console.log('new merge groups');
        //        console.log(JSON.stringify(mg));
        //Colorize.generate_proposed_fixes(mg);
        return mg;
    };
    Colorize.entropy = function (p) {
        return -p * Math.log2(p);
    };
    Colorize.fix_metric = function (target_norm, target, merge_with_norm, merge_with) {
        var n_target = rectangleutils_1.RectangleUtils.area(target);
        var n_merge_with = rectangleutils_1.RectangleUtils.area(merge_with);
        var n_min = Math.min(n_target, n_merge_with);
        var n_max = Math.max(n_target, n_merge_with);
        var norm_min = Math.min(merge_with_norm * n_merge_with, target_norm * n_target);
        var norm_max = Math.max(merge_with_norm * n_merge_with, target_norm * n_target);
        var fix_distance = Math.abs(norm_max - norm_min);
        var entropy_drop = Colorize.entropy(n_min / (n_min + n_max));
        return n_min / (entropy_drop * fix_distance);
    };
    Colorize.generate_proposed_fixes = function (groups) {
        var proposed_fixes = [];
        var already_proposed_pair = {};
        for (var _i = 0, _a = Object.keys(groups); _i < _a.length; _i++) {
            var k1 = _a[_i];
            // Look for possible fixes in OTHER groups.
            for (var i = 0; i < groups[k1].length; i++) {
                var r1 = groups[k1][i];
                var sr1 = JSON.stringify(r1);
                for (var _b = 0, _c = Object.keys(groups); _b < _c.length; _b++) {
                    var k2 = _c[_b];
                    if (k1 === k2) {
                        continue;
                    }
                    for (var j = 0; j < groups[k2].length; j++) {
                        var r2 = groups[k2][j];
                        var sr2 = JSON.stringify(r2);
                        if (!(sr1 + sr2 in already_proposed_pair) && !(sr2 + sr1 in already_proposed_pair)) {
                            if (rectangleutils_1.RectangleUtils.is_mergeable(r1, r2)) {
                                already_proposed_pair[sr1 + sr2] = true;
                                already_proposed_pair[sr2 + sr1] = true;
                                // console.log("could merge (" + k1 + ") " + JSON.stringify(groups[k1][i]) + " and (" + k2 + ") " + JSON.stringify(groups[k2][j]));
                                var metric = Colorize.fix_metric(parseFloat(k1), r1, parseFloat(k2), r2);
                                // was Math.abs(parseFloat(k2) - parseFloat(k1))
                                proposed_fixes.push([metric, r1, r2]);
                            }
                        }
                    }
                }
            }
        }
        // First attribute is the Euclidean norm of the vectors. Differencing corresponds roughly to earth-mover distance.
        // Other attributes are the rectangles themselves. Sort by biggest entropy reduction first, then norm (?).
        proposed_fixes.sort(function (a, b) { return a[0] - b[0]; });
        return proposed_fixes;
    };
    Colorize.merge_groups = function (groups) {
        for (var _i = 0, _a = Object.keys(groups); _i < _a.length; _i++) {
            var k = _a[_i];
            groups[k] = Colorize.merge_individual_groups(JSON.parse(JSON.stringify(groups[k])));
        }
        return groups;
    };
    Colorize.merge_individual_groups = function (group) {
        var numIterations = 0;
        group = group.sort();
        //        console.log(JSON.stringify(group));
        while (true) {
            // console.log("iteration "+numIterations);
            var merged_one = false;
            var deleted_rectangles = {};
            var updated_rectangles = [];
            var working_group = JSON.parse(JSON.stringify(group));
            while (working_group.length > 0) {
                var head = working_group.shift();
                for (var i = 0; i < working_group.length; i++) {
                    //                    console.log("comparing " + head + " and " + working_group[i]);
                    if (rectangleutils_1.RectangleUtils.is_mergeable(head, working_group[i])) {
                        //console.log("friendly!" + head + " -- " + working_group[i]);
                        updated_rectangles.push(rectangleutils_1.RectangleUtils.bounding_box(head, working_group[i]));
                        deleted_rectangles[JSON.stringify(head)] = true;
                        deleted_rectangles[JSON.stringify(working_group[i])] = true;
                        merged_one = true;
                        break;
                    }
                }
                //                if (!merged_one) {
                //                    updated_rectangles.push(head);
                //                }
            }
            for (var i = 0; i < group.length; i++) {
                if (!(JSON.stringify(group[i]) in deleted_rectangles)) {
                    updated_rectangles.push(group[i]);
                }
            }
            updated_rectangles.sort();
            // console.log('updated rectangles = ' + JSON.stringify(updated_rectangles));
            //            console.log('group = ' + JSON.stringify(group));
            if (!merged_one) {
                // console.log('updated rectangles = ' + JSON.stringify(updated_rectangles));
                return updated_rectangles;
            }
            group = JSON.parse(JSON.stringify(updated_rectangles));
            numIterations++;
            if (numIterations > 20) {
                return [[[-1, -1], [-1, -1]]];
            }
        }
    };
    Colorize.generate_all_references = function (formulas, origin_col, origin_row) {
        // Generate all references.
        var refs = {};
        for (var i = 0; i < formulas.length; i++) {
            var row = formulas[i];
            for (var j = 0; j < row.length; j++) {
                // console.log('origin_col = '+origin_col+', origin_row = ' + origin_row);
                var all_deps = excelutils_1.ExcelUtils.all_cell_dependencies(row[j]); // , origin_col + j, origin_row + i);
                if (all_deps.length > 0) {
                    // console.log(all_deps);
                    var src = [origin_col + j, origin_row + i];
                    // console.log('src = ' + src);
                    for (var _i = 0, all_deps_1 = all_deps; _i < all_deps_1.length; _i++) {
                        var dep = all_deps_1[_i];
                        var dep2 = dep; // [dep[0]+origin_col, dep[1]+origin_row];
                        //				console.log('dep type = ' + typeof(dep));
                        //				console.log('dep = '+dep);
                        refs[dep2.join(',')] = refs[dep2.join(',')] || [];
                        refs[dep2.join(',')].push(src);
                        // console.log('refs[' + dep2.join(',') + '] = ' + JSON.stringify(refs[dep2.join(',')]));
                    }
                }
            }
        }
        return refs;
    };
    Colorize.hash_vector = function (vec) {
        // Return a hash of the given vector.
        var h = Math.sqrt(vec.map(function (v) { return v * v; }).reduce(function (a, b) { return a + b; }));
        //	console.log("hash of " + JSON.stringify(vec) + " = " + h);
        return h;
        //        let h = Colorize.hash(JSON.stringify(vec) + 'NONCE01');
        //        return h;
    };
    Colorize.initialized = false;
    Colorize.color_list = [];
    Colorize.light_color_list = [];
    Colorize.light_color_dict = {};
    return Colorize;
}());
exports.Colorize = Colorize;
//console.log(Colorize.dependencies('$C$2:$E$5', 10, 10));
//console.log(Colorize.dependencies('$A$123,A1:B$12,$A12:$B$14', 10, 10));
//console.log(Colorize.hash_vector(Colorize.dependencies('$C$2:$E$5', 10, 10)));