//==============================================================================================================================
// CSV LIBRARY START

/*
 CSV-JS - A Comma-Separated Values parser for JS

 Built to rfc4180 standard, with options for adjusting strictness:
    - optional carriage returns for non-microsoft sources
    - automatically type-cast numeric an boolean values
    - relaxed mode which: ignores blank lines, ignores gargabe following quoted tokens, does not enforce a consistent record length

 Licensed under the MIT license: http://www.opensource.org/licenses/mit-license.php

 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:

 The above copyright notice and this permission notice shall be included in
 all copies or substantial portions of the Software.

 Author Greg Kindel (twitter @gkindel), 2013
 */

    // implemented as a singleton because JS is single threaded
    var CSV = {};
    CSV.RELAXED = false;
    CSV.IGNORE_RECORD_LENGTH = false;
    CSV.IGNORE_QUOTES = false;
    CSV.LINE_FEED_OK = true;
    CSV.CARRIAGE_RETURN_OK = true;
    CSV.DETECT_TYPES = true;
    CSV.IGNORE_QUOTE_WHITESPACE = true;
    CSV.DEBUG = false;

    CSV.ERROR_EOF = "UNEXPECTED_END_OF_FILE";
    CSV.ERROR_CHAR = "UNEXPECTED_CHARACTER";
    CSV.ERROR_EOL = "UNEXPECTED_END_OF_RECORD";
    CSV.WARN_SPACE = "UNEXPECTED_WHITESPACE"; // not per spec, but helps debugging

    var QUOTE = "\"",
        CR = "\r",
        LF = "\n",
        COMMA = ";", // HACK TO SUPPORT BROKEN EXCEL CSV FILES
        SPACE = " ",
        TAB = "\t";

    // states
    var PRE_TOKEN = 0,
        MID_TOKEN = 1,
        POST_TOKEN = 2,
        POST_RECORD = 4;
    /**
     * @name CSV.parse
     * @function
     * @description rfc4180 standard csv parse
     * with options for strictness and data type conversion
     * By default, will automatically type-cast numeric an boolean values.
     * @param {String} str A CSV string
     * @return {Array} An array records, each of which is an array of scalar values.
     * @example
     * // simple
     * var rows = CSV.parse("one,two,three\nfour,five,six")
     * // rows equals [["one","two","three"],["four","five","six"]]
     * @example
     * // Though not a jQuery plugin, it is recommended to use with the $.ajax pipe() method:
     * $.get("csv.txt")
     *    .pipe( CSV.parse )
     *    .done( function(rows) {
     *        for( var i =0; i < rows.length; i++){
     *            console.log(rows[i])
     *        }
     *  });
     * @see http://www.ietf.org/rfc/rfc4180.txt
     */
    CSV.parse = function (str) {
        var result = CSV.result = [];
        CSV.offset = 0;
        CSV.str = str;
        CSV.record_begin();

        CSV.debug("parse()", str);

        var c;
        while( 1 ){
            // pull char
            c = str[CSV.offset++];
            CSV.debug("c", c);

            // detect eof
            if (c == null) {
                if( CSV.escaped )
                    CSV.error(CSV.ERROR_EOF);

                if( CSV.record ){
                    CSV.token_end();
                    CSV.record_end();
                }

                CSV.debug("...bail", c, CSV.state, CSV.record);
                CSV.reset();
                break;
            }

            if( CSV.record == null ){
                // if relaxed mode, ignore blank lines
                if( CSV.RELAXED && (c == LF || c == CR && str[CSV.offset + 1] == LF) ){
                    continue;
                }
                CSV.record_begin();
            }

            // pre-token: look for start of escape sequence
            if (CSV.state == PRE_TOKEN) {

                if ( (c === SPACE || c === TAB) && CSV.next_nonspace() == QUOTE ){
                    if( CSV.RELAXED || CSV.IGNORE_QUOTE_WHITESPACE ) {
                        continue;
                    }
                    else {
                        // not technically an error, but ambiguous and hard to debug otherwise
                        CSV.warn(CSV.WARN_SPACE);
                    }
                }

                if (c == QUOTE && ! CSV.IGNORE_QUOTES) {
                    CSV.debug("...escaped start", c);
                    CSV.escaped = true;
                    CSV.state = MID_TOKEN;
                    continue;
                }
                CSV.state = MID_TOKEN;
            }

            // mid-token and escaped, look for sequences and end quote
            if (CSV.state == MID_TOKEN && CSV.escaped) {
                if (c == QUOTE) {
                    if (str[CSV.offset] == QUOTE) {
                        CSV.debug("...escaped quote", c);
                        CSV.token += QUOTE;
                        CSV.offset++;
                    }
                    else {
                        CSV.debug("...escaped end", c);
                        CSV.escaped = false;
                        CSV.state = POST_TOKEN;
                    }
                }
                else {
                    CSV.token += c;
                    CSV.debug("...escaped add", c, CSV.token);
                }
                continue;
            }

            // fall-through: mid-token or post-token, not escaped
            if (c == CR ) {
                if( str[CSV.offset] == LF  )
                    CSV.offset++;
                else if( ! CSV.CARRIAGE_RETURN_OK )
                    CSV.error(CSV.ERROR_CHAR);
                CSV.token_end();
                CSV.record_end();
            }
            else if (c == LF) {
                if( ! (CSV.LINE_FEED_OK || CSV.RELAXED) )
                    CSV.error(CSV.ERROR_CHAR);
                CSV.token_end();
                CSV.record_end();
            }
            else if (c == COMMA) {
                CSV.token_end();
            }
            else if( CSV.state == MID_TOKEN ){
                CSV.token += c;
                CSV.debug("...add", c, CSV.token);
            }
            else if ( c === SPACE || c === TAB) {
                if (! CSV.IGNORE_QUOTE_WHITESPACE )
                    CSV.error(CSV.WARN_SPACE );
            }
            else if( ! CSV.RELAXED ){
                CSV.error(CSV.ERROR_CHAR);
            }
        }
        return result;
    };

    CSV.reset = function () {
        CSV.state = null;
        CSV.token = null;
        CSV.escaped = null;
        CSV.record = null;
        CSV.offset = null;
        CSV.result = null;
        CSV.str = null;
    };

    CSV.next_nonspace = function () {
        var i = CSV.offset;
        var c;
        while( i < CSV.str.length ) {
            c = CSV.str[i++];
            if( !( c == SPACE || c === TAB ) ){
                return c;
            }
        }
        return null;
    };

    CSV.record_begin = function () {
        CSV.escaped = false;
        CSV.record = [];
        CSV.token_begin();
        CSV.debug("record_begin");
    };

    CSV.record_end = function () {
        CSV.state = POST_RECORD;
        if( ! (CSV.IGNORE_RECORD_LENGTH || CSV.RELAXED)
            && CSV.result.length > 0 && CSV.record.length !=  CSV.result[0].length ){
            CSV.error(CSV.ERROR_EOL);
        }
        CSV.result.push(CSV.record);
        CSV.debug("record end", CSV.record);
        CSV.record = null;
    };

    CSV.resolve_type = function (token) {
        if( token.match(/^\d+(\.\d+)?$/) ){
            token = parseFloat(token);
        }
        else if( token.match(/^true|false$/i) ){
            token = Boolean( token.match(/true/i) );
        }
        else if(token === "undefined" ){
            token = undefined;
        }
        else if(token === "null" ){
            token = null;
        }
        return token;
    };

    CSV.token_begin = function () {
        CSV.state = PRE_TOKEN;
        // considered using array, but http://www.sitepen.com/blog/2008/05/09/string-performance-an-analysis/
        CSV.token = "";
    };

    CSV.token_end = function () {
        if( CSV.DETECT_TYPES ) {
            CSV.token = CSV.resolve_type(CSV.token);
        }
        CSV.record.push(CSV.token);
        CSV.debug("token end", CSV.token);
        CSV.token_begin();
    };

    CSV.debug = function (){
        if( CSV.DEBUG )
            console.log(arguments);
    };

    CSV.dump = function (msg) {
        return [
            msg , "at char", CSV.offset, ":",
            CSV.str.substr(CSV.offset- 50, 50)
                .replace(/\r/mg,"\\r")
                .replace(/\n/mg,"\\n")
                .replace(/\t/mg,"\\t")
        ].join(" ");
    };

    CSV.error = function (err){
        var msg = CSV.dump(err);
        CSV.reset();
        throw msg;
    };

    CSV.warn = function (err){
        var msg = CSV.dump(err);
        try {
            console.warn( msg );
            return;
        } catch (e) {}

        try {
            console.log( msg );
        } catch (e) {}

    };

// CSV LIBRARY END
//==============================================================================================================================
// ОБРАБОТКА ВХОДНЫХ ДАННЫХ - НАЧАЛО

/**
 * Здесь мы конвертируем содержимое файла с данными (CSV файл, разделитель `\t`)
 * в массив записей, каждая запись - массив из шести элементов (по числу колонок).
 * В процессе конвертирования входные данные очищаются от лишних пробелов,
 * корректно сохраняются символы переноса строки, встроенные в значения полей, и т. д.
 **/
convert = function (string) {
	if (!string)
		return [];

	var result = CSV.parse(string, ';');

	if (result[result.length - 1].length == 1 && result[result.length - 1][0] == '')
		result.pop();

	return formatRecords(result);
}

function formatRecords(records) {
	var result = [];
	
	for (var idx = 0; idx < records.length; ++idx) 
		result.push(
			formatFields(
				convertToPersonData(
					normalizeRecordLength(
						records[idx]))));
	
	return result;
}

// Ожидаемое количество колонок во входном CSV
// Фамилия, имя отчество, должность, адрес, телефоны, email, skype 
var NORMAL_RECORD_LENGTH = 7; 

// Убедиться, что в записи ровно столько колонок, сколько нужно
function normalizeRecordLength(record) {
	var diff = NORMAL_RECORD_LENGTH - record.length;
	if (diff < 0) { // record longer
		record = record.slice(0, NORMAL_RECORD_LENGTH);
	} else if (diff > 0) { // record shorter
		for (var i = diff; i > 0; --i) {
			record.push('');
		}
	}

	return record;
}

function convertToPersonData(record) {

	record = cleanRecordFields(record);
	
	var nameParts = getNameParts(record[1]);
	
	return {
		surname:    record[0],
		firstname:       nameParts.firstname,
		fathername: nameParts.fathername,
		duty:       record[2],
		address:    record[3],
		phones:     record[4],
		email:      record[5],
		skype:      record[6],
		website:    ''
	};
}
function cleanRecordFields(record) {
	for (var i = 0; i < record.length; ++i) 
		record[i] = trimString(record[i]);
	
	return record;
}

function getNameParts(namestring) {
	if (!namestring)
		return {firstname: '', fathername: ''};
		
	var nameParts = namestring.split(' ');
	
	if (nameParts.length == 1)
		return {firstname: nameParts[0], fathername: ''};
	
	firstname = nameParts.shift();
	fathername = nameParts.join(' ');
	return {firstname: firstname, fathername: fathername};
}

function formatFields(data) {
	var newdata = data;
	newdata.firstname = capitalizeFirstLetter(data.firstname);
	newdata.fathername = capitalizeFirstLetter(data.fathername);
	
	newdata.duty = correctPunctuation(capitalizeFirstLetter(data.duty));
	
	newdata.address = correctPunctuation(data.address);
	
	newdata.phones = formatPhones(data.phones);
	
	newdata.website = makeWebsiteAddress(data);

	newdata.email = data.email.toLocaleLowerCase();
	
	return newdata;
}

formatPhones = function (string) {
	return string
		.replace(/[\-\+\(]*[0-9][\-\+ \(\)0-9]*[0-9]/g, formatPhone);
}

String.prototype.trim = function () {
    return this
        .replace(/^\s+/, '')
        .replace(/\s+$/, '');
}

String.prototype.collapseSpaces = function () {
    return this
        .replace(/ +/g, ' ');
}

String.prototype.correctSpacesAroundParentheses = function () {
    return this
        .replace(/\s*\(\s*/g, ' (')
        .replace(/\s*\)\s*/g, ') ');
};

formatPhone = function (raw) {
    // If we have short number, there's nothing really what we can do.
    if (raw.length < 10)
        return raw;

    if (raw.match(/[^0-9-+)( ]/))
        return raw;

    // Basic cleanup of the long number
    var base = raw
        .collapseSpaces()
        .correctSpacesAroundParentheses()
        .trim();

    // Trim prefix to go international 
    if (base.charAt(0) == '8' || base.substring(0, 2) == '79')
        base = base.substring(1);
    else if (base.substring(0, 2) == '+7')
        base = base.substring(2);

    // If only numbers then format it as nice as you can.
    if (base.match(/\d+/))
        base = base.replace(/(\d\d\d)(\d\d\d)(\d\d)(\d+)$/, '($1) $2-$3-$4');

    // Replace spaces with dashes
    base = base.replace(/ +/g, '-');

    // Replace zone code inside dashes or spaces with zone code in parentheses
    base = base.replace(/^[- ]*(\d+)[- ]+/, '($1) ');

    // If we have parentheses, GREAT! Remove everything around them and put 8 at start.
    base = base.replace(/.*\((.*)\)[^0-9]+/, '8 ($1) ');

    // If zone code starts with 9, it's mobile number. Start it with international +7 
    // instead of Russia-specific 8
    base = base.replace(/^8 \(9/, '+7 (9');

    return base;
}

function makeWebsiteAddress(data) {
	var website = 'www.trakt.ru';

	var subpath = extractCityId(data.address);
	
	if (!subpath)
		subpath = extractCityId(data.duty);
		
	if (subpath)
		website = website + '/' + subpath;

	return website;
}

function extractCityId(str) {
	var cities = [
		['Архангельск', 'arkhangelsk'],
		['Астрахань', 'astrakhan'],
		['Балаково', 'balakovo'],
		['Благовещенск', 'blagoveshensk'],
		['Владивосток', 'vladivostok'],
		['Владимир', 'vladimir'],
		['Волгоград', 'volgograd'],
		['Волжский', 'volzhsky'],
		['Воронеж', 'voronej'],
		['Екатеринбург', 'ekaterinburg'],
		['Ижевск', 'ijevsk'],
		['Иркутск', 'irkutsk'],
		['Казань', 'kazan'],
		['Калуга', 'kaluga'],
		['Киров', 'kirov'],
		['Кострома', 'kostroma'],
		['Краснодар', 'krasnodar'],
		['Красноярск', 'krasnoyarsk'],
		['Липецк', 'lipetsk'],
		['Миасс', 'miass'],
        // Москва была удалена, чтобы URL выглядел просто www.trakt.ru
		['Мурманск', 'murmansk'],
		['Набережные Челны', 'chelny'],
		['Нижневартовск', 'nizhnevartovsk'],
		['Нижний Новгород', 'nnovgorod'],
		['Новокузнецк', 'novokuzneck'],
		['Новороссийск', 'novorossiysk'],
		['Новосибирск', 'novosibirsk'],
		['Омск', 'omsk'],
		['Оренбург', 'orenburg'],
		['Орёл', 'orel'],
		['Пермь', 'perm'],
		['Петрозаводск', 'petrozavodsk'],
		['Ростов-на-Дону', 'rostov'],
		['Рязань', 'ryazan'],
		['Самара', 'samara'],
		['Санкт-Петербург', 'peterburg'],
		['Саратов', 'saratov'],
		['Смоленск', 'smolensk'],
		['Сочи', 'sochi'],
		['Ставрополь ', 'stavropol'],
		['Сургут', 'surgut'],
		['Сыктывкар', 'syktyvkar'],
		['Тверь', 'tver'],
		['Тольятти', 'togliatti'],
		['Тула', 'tula'],
		['Тюмень', 'tumen'],
		['Ульяновск', 'ulyanovsk'],
		['Уфа', 'ufa'],
		['Челябинск', 'chelyabinsk'],
		['Череповец', 'cherepovets'],
		['Ярославль', 'yaroslavl']
	];
	for (var idx = 0; idx < cities.length; ++idx) {
		var cityname = cities[idx][0],
			cityid = cities[idx][1],
			found = str.search(new RegExp(cityname, 'i'));
		
		if (found > -1 && cityname !== 'Москва') 
			return cityid;
	}
	return '';
}

function correctPunctuation(str) {
	return str
		.replace(/([а-яА-Я0-9])\s*([,.;:])\s*/g, '$1$2 ')
		.replace(/\s\s*$/, ''); // right trim
}

function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}

function trimString(str) {
	return str
		.replace(/^\s\s*/, '') // left trim
		.replace(/\s\s*$/, '') // right trim
		.replace(/\r?\n|\r/g, ' ') // replace line breaks with spaces
		.replace(/ {2,}/g, ' '); // collapse inner spaces
}

// ОБРАБОТКА ВХОДНЫХ ДАННЫХ - КОНЕЦ
//========================================================================================================
// РАБОТА С САМИМ МАКЕТОМ - НАЧАЛО

// Главная функция для того, чтобы можно было сделать guard case'ы
function main() {
    if (app.documents.length == 0) {
      alert("Не открыт шаблон, в который нужно вставить данные - прекращаем работу.");
      return;
   }

    // Откуда брать данные
    var datafile = File.openDialog ("Выберите файл с данными", "Формат CSV (разделитель - точка с запятой):*.csv", false);
    if (!datafile) {
        alert ('Не выбран файл с данными - прекращаем работу.');
        return;
    }

    // Куда сохранять
    var destFolder = Folder.selectDialog( 'Куда сохранять готовые файлы?', '~' );
    if (!destFolder) {
        alert ('Не указана папка для сохранения файлов - прекращаем работу.');
        return;
    }

    datafile.open('r');
    var raw = datafile.read();

    var data = convert(raw);

    // Открываем макет (это должен быть макет визитки)
    var doc = activeDocument;

    var originalFilename = '';
    for (var idx = 0; idx < data.length; ++idx) {
        var record = data[idx];
        if (!record)
            continue;

        insertDataInDoc(record, doc);
        makeAdditionalFormatting(doc);

        // Файл для сохранения
        var filename = makeFilename(destFolder, record);
        originalFilename = doc.fullName;

        saveDocAsEPS (doc, filename);

        if (idx > 0) { // after "save as" we have the same document opened by new name. If we do "app.open" on the new name Illustrator will do nothing, as this file is opened already.
            app.open(originalFilename);
        }

        app.redraw();
    }
    alert ('Всё должно быть выполнено.');
}

function makeAdditionalFormatting(doc) {
    setRightJustification(doc.pageItems.getByName('Контакты'));
}

function setRightJustification(contactsBlock) {
        contactsBlock.textRange.paragraphAttributes.justification = Justification.RIGHT;
        app.redraw();
 }

function makeFilename(foldername, record) {
    var filename = record.surname;
    return foldername + '/' + filename;
}

/**
 * Данные должны быть набором строк следующего вида:
 * Фамилия ИмяОтчество  Должность   Адрес   Телефоны    E-mail  Skype
 */
function insertDataInDoc(data, doc) {
    doc.pageItems.getByName('ФИО').contents = data.surname.toLocaleUpperCase() + "\r" + data.firstname + ' ' + data.fathername; 
    doc.pageItems.getByName('Должность').contents = data.duty;
    doc.pageItems.getByName('Адрес').contents = data.address + "\r" + data.phones; 
    if (data.skype)
        data.skype = '\rSkype: ' + data.skype;
    doc.pageItems.getByName('Контакты').contents = data.website + "\r" + data.email + data.skype;
}

function saveDocAsEPS(doc, filename) {
    var epsFile = new File(filename + '.eps');
    var epsOptions = new EPSSaveOptions();
    doc.saveAs(epsFile, epsOptions);
}

// РАБОТА С САМИМ МАКЕТОМ - КОНЕЦ
//========================================================================================================


main();
