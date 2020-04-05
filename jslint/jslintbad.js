// jslint.js
// 2007-01-14
// Special version for use only with JavaScript Plus!
// By Luis Nunez
// VBSoftware
// http://www.vbsoftware.cl

/*
Copyright (c) 2002 Douglas Crockford  (www.JSLint.com)

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the &quot;Software&quot;), to
deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
of the Software, and to permit persons to whom the Software is furnished to do
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

The Software shall be used for Good, not Evil.

THE SOFTWARE IS PROVIDED &quot;AS IS&quot;, WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/


/*
    JSLINT is a global function. It takes two parameters.

        var myResult = JSLINT(source, option);

    The first parameter is either a string or an array of strings. If it is a
    string, it will be split on &#39;\n&#39; or &#39;\r&#39;. If it is an array
of strings, it
    is assumed that each string represents one line. The source can be a
    JavaScript text, or HTML text, or a Konfabulator text.

    The second parameter is an optional object of options which control the
    operation of JSLINT. All of the options are booleans. All are optional and
    have a default value of false.

    {
        browser    : true if the standard browser globals should be predefined
        cap        : true if upper case HTML should be allowed
        debug      : true if debugger statements should be allowed
        eqeqeq     : true if === should be required
        evil       : true if eval should be allowed
        jscript    : true if jscript deviations should be allowed
        laxLineEnd : true if line breaks should not be checked
        passfail   : true if the scan should stop on first error
        plusplus   : true if increment/decrement should not be allowed
        redef      : true if var redefinition should be allowed
        undef      : true if undefined variables are errors
        widget     : true if the Yahoo Widgets globals should be predefined
    }

    If it checks out, JSLINT returns true. Otherwise, it returns false.

    If false, you can inspect JSLINT.errors to find out the problems.
    JSLINT.errors is an array of objects containing these members:

    {
        line      : The line (relative to 0) at which the lint was found
        character : The character (relative to 0) at which the lint was found
        reason    : The problem
        evidence  : The text line in which the problem occurred
    }

    If a fatal error was found, a null will be the last element of the
    JSLINT.errors array.

    You can request a Function Report, which shows all of the functions
    and the parameters and vars that they use. This can be used to find
    implied global variables and other problems. The report is in HTML and
    can be inserted in a &lt;body&gt;.

        var myReport = JSLINT.report(option);

    If the option is true, then the report will be limited to only errors.
*/

String.prototype.entityify = function () {
    return this.
        replace(/&amp;/g, &#39;&amp;amp;&#39;).
        replace(/&lt;/g, &#39;&amp;lt;&#39;).
        replace(/&gt;/g, &#39;&amp;gt;&#39;);
};

String.prototype.isAlpha = function () {
    return (this &gt;= &#39;a&#39; &amp;&amp; this &lt;= &#39;z\uffff&#39;) ||
        (this &gt;= &#39;A&#39; &amp;&amp; this &lt;= &#39;Z\uffff&#39;);
};


String.prototype.isDigit = function () {
    return (this &gt;= &#39;0&#39; &amp;&amp; this &lt;= &#39;9&#39;);
};


// We build the application inside a function so that we produce only a single
// global variable. The function will be invoked, its return value is the
JSLINT
// function itself.

var JSLINT;
JSLINT = function () {

    var anonname,

// browser contains a set of global names which are commonly provided by a
// web browser environment.

        browser = {
            alert: true,
            blur: true,
            clearInterval: true,
            clearTimeout: true,
            close: true,
            closed: true,
            confirm: true,
            defaultStatus: true,
            document: true,
            event: true,
            focus: true,
            frames: true,
            history: true,
            Image: true,
            length: true,
            location: true,
            moveBy: true,
            moveTo: true,
            name: true,
            navigator: true,
            onblur: true,
            onerror: true,
            onfocus: true,
            onload: true,
            onresize: true,
            onunload: true,
            open: true,
            opener: true,
            parent: true,
            print: true,
            prompt: true,
            resizeBy: true,
            resizeTo: true,
            screen: true,
            scroll: true,
            scrollBy: true,
            scrollTo: true,
            self: true,
            setInterval: true,
            setTimeout: true,
            status: true,
            top: true,
            window: true,
            XMLHttpRequest: true
        },
        funlab, funstack, functions, globals,

// konfab contains the global names which are provided to a Yahoo
// (fna Konfabulator) widget.

        konfab = {
            alert: true,
            animator: true,
            appleScript: true,
            beep: true,
            bytesToUIString: true,
            chooseColor: true,
            chooseFile: true,
            chooseFolder: true,
            convertPathToHFS: true,
            convertPathToPlatform: true,
            closeWidget: true,
            CustomAnimation: true,
            escape: true,
            FadeAnimation: true,
            focusWidget: true,
            form: true,
            include: true,
            isApplicationRunning: true,
            iTunes: true,
            konfabulatorVersion: true,
            log: true,
            MoveAnimation: true,
            openURL: true,
            play: true,
            popupMenu: true,
            print: true,
            prompt: true,
            reloadWidget: true,
            resolvePath: true,
            resumeUpdates: true,
            RotateAnimation: true,
            runCommand: true,
            runCommandInBg: true,
            saveAs: true,
            savePreferences: true,
            showWidgetPreferences: true,
            sleep: true,
            speak: true,
            suppressUpdates: true,
            tellWidget: true,
            unescape: true,
            updateNow: true,
            yahooCheckLogin: true,
            yahooLogin: true,
            yahooLogout: true,
            COM: true,
            filesystem: true,
            preferenceGroups: true,
            preferences: true,
            screen: true,
            system: true,
            URL: true,
            XMLDOM: true,
            XMLHttpRequest: true
        },
        lines, lookahead, member, noreach, option, prevtoken, stack,

// standard contains the global names that are provided by standard JavaScript.

        standard = {
            Array: true,
            Boolean: true,
            Date: true,
            decodeURI: true,
            decodeURIComponent: true,
            encodeURI: true,
            encodeURIComponent: true,
            Error: true,
            escape: true,
            &#39;eval&#39;: true,
            EvalError: true,
            Function: true,
            isFinite: true,
            isNaN: true,
            Math: true,
            Number: true,
            Object: true,
            parseInt: true,
            parseFloat: true,
            RangeError: true,
            ReferenceError: true,
            RegExp: true,
            String: true,
            SyntaxError: true,
            TypeError: true,
            unescape: true,
            URIError: true
        },
        syntax = {}, token, verb,

//  xmode is used to adapt to the exceptions in XML parsing.
//  It can have these states:
//      false   .js script file
//      &quot;       A &quot; attribute
//      &#39;       A &#39; attribute
//      content The content of a script tag
//      CDATA   A CDATA block

        xmode,

//  xtype identifies the type of document being analyzed.
//  It can have these states:
//      false   .js script file
//      html    .html file
//      widget  .kon Konfabulator file

        xtype,
// token
        tx =
/^([(){}[.,:;&#39;&quot;~]|\](\]&gt;)?|\?&gt;?|==?=?|\/(\*(global|extern)*|=|)|\*[\/=]?|\+[+=]?|-[-=]?|%[=&gt;]?|&amp;[&amp;=]?|\|[|=]?|&gt;&gt;?&gt;?=?|&lt;([\/=%\?]|\!(\[|--)?|&lt;=?)?|\^=?|\!=?=?|[a-zA-Z_$][a-zA-Z0-9_$]*|[0-9]+([xX][0-9a-fA-F]+|\.[0-9]*)?([eE][+-]?[0-9]+)?)/,
// string ending in single quote
        sx = /^((\\[^\x00-\x1f]|[^\x00-\x1f&#39;\\])*)&#39;/,
        sxx = /^(([^\x00-\x1f&#39;])*)&#39;/,
// string ending in double quote
        qx = /^((\\[^\x00-\x1f]|[^\x00-\x1f&quot;\\])*)&quot;/,
        qxx = /^(([^\x00-\x1f&quot;])*)&quot;/,
// regular expression
        rx =
/^(\\[^\x00-\x1f]|\[(\\[^\x00-\x1f]|[^\x00-\x1f\\\/])*\]|[^\x00-\x1f\\\/\[])+\/[gim]*/,
// star slash
        lx = /\*\/|\/\*/,
// global identifier
        gx = /^([a-zA-Z_$][a-zA-Z0-9_$]*)/,
// identifier
        ix = /^([a-zA-Z_$][a-zA-Z0-9_$]*$)/,
// global separators
        hx = /^[\x00-\x20,]*(\*\/)?/,
// whitespace
        wx = /^\s*(\/\/.*\r*$)?/;

// Make a new object that inherits from an existing object.

    function object(o) {
        function F() {}
        F.prototype = o;
        return new F();
    }

// Produce an error warning.

    function warning(m, x, y) {
        var l, c, t = typeof x === &#39;object&#39; ? x : token;
        if (typeof x === &#39;number&#39;) {
            l = x;
            c = y || 0;
        } else {
            if (t.id === &#39;(end)&#39;) {
                t = prevtoken;
            }
            l = t.line || 0;
            c = t.from || 0;
        }
        JSLINT.errors.push({
            id: &#39;(error)&#39;,
            reason: m,
            evidence: lines[l] || &#39;&#39;,
            line: l,
            character: c
        });
        if (option.passfail) {
            JSLINT.errors.push(null);
            throw null;
        }
    }

    function error(m, x, y) {
        warning(m, x, y);
        warning(&quot;Stopping, unable to continue.&quot;, x, y);
        JSLINT.errors.push(null);
        throw null;
    }


// lexical analysis

    var lex = function () {
        var character, from, line, s;

// Private lex methods

        function nextLine() {
            line += 1;
            if (line &gt;= lines.length) {
                return false;
            }
            character = 0;
            s = lines[line];
            return true;
        }

// Produce a token object.  The token inherits from a syntax symbol.

        function it(type, value) {
            var t;
            if (type === &#39;(punctuator)&#39;) {
                t = syntax[value];
            } else if (type === &#39;(identifier)&#39;) {
                t = syntax[value];
                if (!t || typeof t !== &#39;object&#39;) {
                    t = syntax[type];
                }
            } else {
                t = syntax[type];
            }
            if (!t || typeof t !== &#39;object&#39;) {
                error(&quot;Unrecognized symbol: &#39;&quot; + value +
&quot;&#39; &quot; + type);
            }
            t = object(t);
            if (value || type === &#39;(string)&#39;) {
                if (value.charAt(10) === &#39;:&#39; &amp;&amp;
                        value.substring(0, 10).toLowerCase() ===
&#39;javascript&#39;) {
                    warning(&quot;JavaScript URL.&quot;);
                }
                t.value = value;
            }
            t.line = line;
            t.character = character;
            t.from = from;
            return t;
        }

// Public lex methods

        return {
            init: function (source) {
                if (typeof source === &#39;string&#39;) {
                    lines = source.split(&#39;\n&#39;);
                    if (lines.length === 1) {
                        lines = lines[0].split(&#39;\r&#39;);
                    }
                } else {
                    lines = source;
                }
                line = 0;
                character = 0;
                from = 0;
                s = lines[0];
            },

// token -- this is called by advance to get the next token.

            token: function () {
                var c, i, l, r, t;

                function string(x) {
                    var a, j;
                    r = x.exec(s);
                    if (r) {
                        t = r[1];
                        l = r[0].length;
                        s = s.substr(l);
                        character += l;
                        if (xmode === &#39;script&#39;) {
                            if (t.indexOf(&#39;&lt;\/&#39;) &gt;= 0) {
                                warning(
    &#39;Expected &quot;...&lt;\\/...&quot; and instead saw
&quot;...&lt;\/...&quot;.&#39;, token);
                            }
                        }
                        return it(&#39;(string)&#39;, r[1]);
                    } else {
                        for (j = 0; j &lt; s.length; j += 1) {
                            a = s.charAt(j);
                            if (a &lt; &#39; &#39;) {
                                if (a === &#39;\n&#39; || a === &#39;\r&#39;) {
                                    break;
                                }
                                warning(&quot;Control character in string:
&quot; +
                                        s.substring(0, j), line, character +
j);
                            }
                        }
                        error(&quot;Unclosed string: &quot; + s, line,
character);
                    }
                }

                for (;;) {
                    if (!s) {
                        return it(nextLine() ? &#39;(endline)&#39; :
&#39;(end)&#39;, &#39;&#39;);
                    }
                    r = wx.exec(s);
                    if (!r || !r[0]) {
                        break;
                    }
                    l = r[0].length;
                    s = s.substr(l);
                    character += l;
                    if (s) {
                        break;
                    }
                }
                from = character;
                r = tx.exec(s);
                if (r) {
                    t = r[0];
                    l = t.length;
                    s = s.substr(l);
                    character += l;
                    c = t.substr(0, 1);

//      identifier

                    if (c.isAlpha() || c === &#39;_&#39; || c === &#39;$&#39;)
{
                        return it(&#39;(identifier)&#39;, t);
                    }

//      number

                    if (c.isDigit()) {
                        if (token.id === &#39;.&#39;) {
                            warning(
            &quot;A decimal fraction should have a zero before the decimal
point.&quot;,
                                token);
                        }
                        if (!isFinite(Number(t))) {
                            warning(&quot;Bad number: &#39;&quot; + t +
&quot;&#39;.&quot;,
                                line, character);
                        }
                        if (s.substr(0, 1).isAlpha()) {
                            warning(&quot;Space is required after a number:
&#39;&quot; +
                                    t + &quot;&#39;.&quot;, line, character);
                        }
                        if (c === &#39;0&#39; &amp;&amp;
t.substr(1,1).isDigit()) {
                            warning(&quot;Don&#39;t use extra leading zeros:
&#39;&quot; +
                                    t + &quot;&#39;.&quot;, line, character);
                        }
                        if (t.substr(t.length - 1) === &#39;.&#39;) {
                            warning(
    &quot;A trailing decimal point can be confused with a dot: &#39;&quot; + t
+ &quot;&#39;.&quot;,
                                    line, character);
                        }
                        return it(&#39;(number)&#39;, t);
                    }

//      string

                    if (t === &#39;&quot;&#39;) {
                        return (xmode === &#39;&quot;&#39; ||  xmode ===
&#39;string&#39;) ?
                            it(&#39;(punctuator)&#39;, t) :
                            string(xmode === &#39;xml&#39; ? qxx : qx);
                    }
                    if (t === &quot;&#39;&quot;) {
                        return (xmode === &quot;&#39;&quot; ||  xmode ===
&#39;string&#39;) ?
                            it(&#39;(punctuator)&#39;, t) :
                            string(xmode === &#39;xml&#39; ? sxx : sx);
                    }

//      unbegun comment

                    if (t === &#39;/*&#39;) {
                        for (;;) {
                            i = s.search(lx);
                            if (i &gt;= 0) {
                                break;
                            }
                            if (!nextLine()) {
                                error(&quot;Unclosed comment.&quot;, token);
                            }
                        }
                        character += i + 2;
                        if (s.substr(i, 1) === &#39;/&#39;) {
                            error(&quot;Nested comment.&quot;);
                        }
                        s = s.substr(i + 2);
                        return this.token();
                    }

//      /*extern

                    if (t === &#39;/*extern&#39; || t === &#39;/*global&#39;) {
                        for (;;) {
                            r = hx.exec(s);
                            if (r) {
                                l = r[0].length;
                                s = s.substr(l);
                                character += l;
                                if (r[1] === &#39;*/&#39;) {
                                    return this.token();
                                }
                            }
                            if (s) {
                                r = gx.exec(s);
                                if (r) {
                                    l = r[0].length;
                                    s = s.substr(l);
                                    character += l;
                                    globals[r[1]] = true;
                                } else {
                                    error(&quot;Bad extern identifier:
&#39;&quot; +
                                        s + &quot;&#39;.&quot;, line,
character);
                                }
                             } else if (!nextLine()) {
                                error(&quot;Unclosed comment.&quot;);
                            }
                        }
                    }

//      punctuator

                    return it(&#39;(punctuator)&#39;, t);
                }
                error(&quot;Unexpected token: &quot; + (t || s.substr(0, 1)),
                    line, character);
            },

// skip -- skip past the next occurrence of a particular string.
// If the argument is empty, skip to just before the next &#39;&lt;&#39;
character.
// This is used to ignore HTML content. Return false if it isn&#39;t found.

            skip: function (to) {
                if (token.id) {
                    if (!to) {
                        to = &#39;&#39;;
                        if (token.id.substr(0, 1) === &#39;&lt;&#39;) {
                            lookahead.push(token);
                            return true;
                        }
                    } else if (token.id.indexOf(to) &gt;= 0) {
                        return true;
                    }
                }
                prevtoken = token;
                token = syntax[&#39;(error)&#39;];
                for (;;) {
                    var i = s.indexOf(to || &#39;&lt;&#39;);
                    if (i &gt;= 0) {
                        character += i + to.length;
                        s = s.substr(i + to.length);
                        return true;
                    }
                    if (!nextLine()) {
                        break;
                    }
                }
                return false;
            },

// regex -- this is called by parse when it sees &#39;/&#39; being used as a
prefix.

            regex: function () {
                var l, r = rx.exec(s), x;
                if (r) {
                    l = r[0].length;
                    character += l;
                    s = s.substr(l);
                    x = r[1];
                    return it(&#39;(regex)&#39;, x);
                }
                error(&quot;Bad regular expression: &quot; + s);
            }
        };
    }();

    function builtin(name) {
        return standard[name] === true ||
               globals[name] === true ||
             ((xtype === &#39;widget&#39; || option.widget) &amp;&amp;
konfab[name] === true) ||
             ((xtype === &#39;html&#39; || option.browser) &amp;&amp;
browser[name] === true);
    }

    function addlabel(t, type) {
        if (t) {
            if (typeof funlab[t] === &#39;string&#39;) {
                switch (funlab[t]) {
                case &#39;var&#39;:
                case &#39;var*&#39;:
                    if (type === &#39;global&#39;) {
                        funlab[t] = &#39;var*&#39;;
                        return;
                    }
                    break;
                case &#39;global&#39;:
                    if (type === &#39;var&#39;) {
                        warning(&#39;Var &#39; + t +
                            &#39; was used before it was declared.&#39;,
prevtoken);
                        return;
                    }
                    if (type === &#39;var*&#39; || type === &#39;global&#39;) {
                        return;
                    }
                    break;
                case &#39;function&#39;:
                case &#39;parameter&#39;:
                    if (type === &#39;global&#39;) {
                        return;
                    }
                    break;
                }
                warning(&quot;Identifier &#39;&quot; + t + &quot;&#39; already
declared as &quot; +
                        funlab[t], prevtoken);
            }
            funlab[t] = type;
        }
    }


// We need a peek function. If it has an argument, it peeks that much farther
// ahead. It is used to distinguish
//     for ( var i in ...
// from
//     for ( var i = ...

    function peek(i) {
        var j = 0, t;
        if (token === syntax[&#39;(error)&#39;]) {
            return token;
        }
        if (typeof i === &#39;undefined&#39;) {
            i = 0;
        }
        while (j &lt;= i) {
            t = lookahead[j];
            if (!t) {
                t = lookahead[j] = lex.token();
            }
            j += 1;
        }
        return t;
    }


    var badbreak = {&#39;)&#39;: true, &#39;]&#39;: true, &#39;++&#39;: true,
&#39;--&#39;: true};

// Produce the next token. It looks for programming errors.

    function advance(id, t) {
        var l;
        switch (prevtoken.id) {
        case &#39;(number)&#39;:
            if (token.id === &#39;.&#39;) {
                warning(
&quot;A dot following a number can be confused with a decimal point.&quot;,
prevtoken);
            }
            break;
        case &#39;-&#39;:
            if (token.id === &#39;-&#39; || token.id === &#39;--&#39;) {
                warning(&quot;Confusing minusses.&quot;);
            }
            break;
        case &#39;+&#39;:
            if (token.id === &#39;+&#39; || token.id === &#39;++&#39;) {
                warning(&quot;Confusing plusses.&quot;);
            }
            break;
        }
        if (prevtoken.type === &#39;(string)&#39; || prevtoken.identifier) {
            anonname = prevtoken.value;
        }

        if (id &amp;&amp; token.value !== id) {
            if (t) {
                if (token.id === &#39;(end)&#39;) {
                    warning(&quot;Unmatched &#39;&quot; + t.id +
&quot;&#39;.&quot;, t);
                } else {
                    warning(&quot;Expected &#39;&quot; + id + &quot;&#39; to
match &#39;&quot; +
                            t.id + &quot;&#39; from line &quot; + (t.line + 1)
+
                            &quot; and instead saw &#39;&quot; + token.value +
&quot;&#39;.&quot;);
                }
            } else {
                warning(&quot;Expected &#39;&quot; + id + &quot;&#39; and
instead saw &#39;&quot; +
                        token.value + &quot;&#39;.&quot;);
            }
        }
        prevtoken = token;
        for (;;) {
            token = lookahead.shift() || lex.token();
            if (token.id === &#39;&lt;![&#39;) {
                if (xtype === &#39;html&#39;) {
                    error(&quot;Unexpected token &#39;&lt;![&#39;&quot;);
                }
                if (xmode === &#39;script&#39;) {
                    token = lex.token();
                    if (token.value !== &#39;CDATA&#39;) {
                        error(&quot;Expected &#39;CDATA&#39;&quot;);
                    }
                    token = lex.token();
                    if (token.id !== &#39;[&#39;) {
                        error(&quot;Expected &#39;[&#39;&quot;);
                    }
                    xmode = &#39;CDATA&#39;;
                } else if (xmode === &#39;xml&#39;) {
                    lex.skip(&#39;]]&gt;&#39;);
                } else {
                    error(&quot;Unexpected token &#39;&lt;![&#39;&quot;);
                }
            } else if (token.id === &#39;]]&gt;&#39;) {
                if (xmode === &#39;CDATA&#39;) {
                    xmode = &#39;script&#39;;
                } else {
                    error(&quot;Unexpected token &#39;]]&gt;&quot;);
                }
            } else if (token.id !== &#39;(endline)&#39;) {
                break;
            }
            if (xmode === &#39;&quot;&#39; || xmode === &quot;&#39;&quot;) {
                error(&quot;Missing &#39;&quot; + xmode + &quot;&#39;.&quot;,
prevtoken);
            }
            l = !xmode &amp;&amp; !option.laxLineEnd &amp;&amp;
                (prevtoken.type === &#39;(string)&#39; || prevtoken.type ===
&#39;(number)&#39; ||
                prevtoken.type === &#39;(identifier)&#39; ||
badbreak[prevtoken.id]);
        }
        if (l) {
            switch (token.id) {
            case &#39;{&#39;:
            case &#39;}&#39;:
            case &#39;]&#39;:
                break;
            case &#39;)&#39;:
                switch (prevtoken.id) {
                case &#39;)&#39;:
                case &#39;}&#39;:
                case &#39;]&#39;:
                    break;
                default:
                    warning(&quot;Line breaking error: &#39;)&#39;.&quot;,
prevtoken);
                }
                break;
            default:
                warning(&quot;Line breaking error: &#39;&quot; +
prevtoken.value + &quot;&#39;.&quot;,
                        prevtoken);
            }
        }
        if (xtype === &#39;widget&#39; &amp;&amp; xmode === &#39;script&#39;
&amp;&amp; token.id) {
            l = token.id.charAt(0);
            if (l === &#39;&lt;&#39; || l === &#39;&amp;&#39;) {
                token.nud = token.led = null;
                token.lbp = 0;
                token.reach = true;
            }
        }
    }


    function advanceregex() {
        token = lex.regex();
    }


    function beginfunction(i) {
        var f = {&#39;(name)&#39;: i, &#39;(line)&#39;: token.line + 1,
&#39;(context)&#39;: funlab};
        funstack.push(funlab);
        funlab = f;
        functions.push(funlab);
    }


    function endfunction() {
        funlab = funstack.pop();
    }


// This is the heart of JSLINT, the Pratt parser. In addition to parsing, it
// is looking for ad hoc lint patterns. We add to Pratt&#39;s model .fud, which
is
// like nud except that it is only used on the first token of a statement.
// Having .fud makes it much easier to define JavaScript. I retained
Pratt&#39;s
// nomenclature, even though it isn&#39;t very descriptive.

// .nud     Null denotation
// .fud     First null denotation
// .led     Left denotation
//  lbp     Left binding power
//  rbp     Right binding power

// They are key to the parsing method called Top Down Operator Precedence.

    function parse(rbp, initial) {
        var l, left, o;
        if (token.id &amp;&amp; token.id === &#39;/&#39;) {
            if (prevtoken.id !== &#39;(&#39; &amp;&amp; prevtoken.id !==
&#39;=&#39; &amp;&amp;
                    prevtoken.id !== &#39;:&#39; &amp;&amp; prevtoken.id !==
&#39;,&#39; &amp;&amp;
                    prevtoken.id !== &#39;=&#39; &amp;&amp; prevtoken.id !==
&#39;[&#39;) {
                warning(
&quot;Expected to see a &#39;(&#39; or &#39;=&#39; or &#39;:&#39; or
&#39;,&#39; or &#39;[&#39; preceding a regular expression literal, and instead
saw &#39;&quot; +
                        prevtoken.value + &quot;&#39;.&quot;, prevtoken);
            }
            advanceregex();
        }
        if (token.id === &#39;(end)&#39;) {
            warning(&quot;Unexpected early end of program&quot;, prevtoken);
        }
        advance();
        if (initial) {
            anonname = &#39;anonymous&#39;;
            verb = prevtoken.value;
        }
        if (initial &amp;&amp; prevtoken.fud) {
            prevtoken.fud();
        } else {
            if (prevtoken.nud) {
                o = prevtoken.exps;
                left = prevtoken.nud();
            } else {
                if (token.type === &#39;(number)&#39; &amp;&amp; prevtoken.id
=== &#39;.&#39;) {
                    warning(
&quot;A leading decimal point can be confused with a dot: .&quot; +
token.value,
                            prevtoken);
                }
                error(&quot;Expected an identifier and instead saw &#39;&quot;
+
                        prevtoken.id + &quot;&#39;.&quot;, prevtoken);
            }
            while (rbp &lt; token.lbp) {
                o = token.exps;
                advance();
                if (prevtoken.led) {
                    left = prevtoken.led(left);
                } else {
                    error(&quot;Expected an operator and instead saw
&#39;&quot; +
                        prevtoken.id + &quot;&#39;.&quot;);
                }
            }
            if (initial &amp;&amp; !o) {
                warning(
&quot;Expected an assignment or function call and instead saw an
expression.&quot;,
                        prevtoken);
            }
        }
        if (l) {
            funlab[l] = &#39;label&#39;;
        }
        if (left &amp;&amp; left.id === &#39;eval&#39;) {
            warning(&quot;evalError&quot;, left);
        }
        return left;
    }


// Parasitic constructors for making the symbols that will be inherited by
// tokens.

    function symbol(s, p) {
        return syntax[s] || (syntax[s] = {id: s, lbp: p, value: s});
    }


    function delim(s) {
        return symbol(s, 0);
    }


    function stmt(s, f) {
        var x = delim(s);
        x.identifier = x.reserved = true;
        x.fud = f;
        return x;
    }


    function blockstmt(s, f) {
        var x = stmt(s, f);
        x.block = true;
        return x;
    }


    function prefix(s, f) {
        var x = symbol(s, 150);
        x.nud = (typeof f === &#39;function&#39;) ? f : function () {
            if (option.plusplus &amp;&amp; (this.id === &#39;++&#39; || this.id
=== &#39;--&#39;)) {
                warning(this.id + &quot; is considered harmful.&quot;, this);
            }
            parse(150);
            return this;
        };
        return x;
    }


    function prefixname(s, f) {
        var x = prefix(s, f);
        x.identifier = x.reserved = true;
        return x;
    }


    function type(s, f) {
        var x = delim(s);
        x.type = s;
        x.nud = f;
        return x;
    }


    function reserve(s, f) {
        var x = type(s, f);
        x.identifier = x.reserved = true;
        return x;
    }


    function reservevar(s) {
        return reserve(s, function () {
            return this;
        });
    }


    function infix(s, f, p) {
        var x = symbol(s, p);
        x.led = (typeof f === &#39;function&#39;) ? f : function (left) {
            return [f, left, parse(p)];
        };
        return x;
    }


    function assignop(s, f) {
        symbol(s, 20).exps = true;
        return infix(s, function (left) {
            if (left) {
                if (left.id === &#39;.&#39; || left.id === &#39;[&#39; ||
                        (left.identifier &amp;&amp; !left.reserved)) {
                    parse(19);
                    return left;
                }
                if (left === syntax[&#39;function&#39;]) {
                    if (option.jscript) {
                        parse(19);
                        return left;
                    } else {
                        warning(
&quot;Expected an identifier in an assignment, and instead saw a function
invocation.&quot;,
                                prevtoken);
                    }
                }
            }
            error(&quot;Bad assignment.&quot;, this);
        }, 20);
    }


    function suffix(s, f) {
        var x = symbol(s, 150);
        x.led = function (left) {
            if (option.plusplus) {
                warning(this.id + &quot; is considered harmful.&quot;, this);
            }
            return [f, left];
        };
        return x;
    }


    function optionalidentifier() {
        if (token.reserved) {
            warning(&quot;Expected an identifier and instead saw &#39;&quot; +
                token.id + &quot;&#39; (a reserved word).&quot;);
        }
        if (token.identifier) {
            advance();
            return prevtoken.value;
        }
    }


    function identifier() {
        var i = optionalidentifier();
        if (i) {
            return i;
        }
        if (prevtoken.id === &#39;function&#39; &amp;&amp; token.id ===
&#39;(&#39;) {
            warning(&quot;Missing name in function statement.&quot;);
        } else {
            error(&quot;Expected an identifier and instead saw &#39;&quot; +
                    token.value + &quot;&#39;.&quot;, token);
        }
    }


    function reachable(s) {
        var i = 0, t;
        if (token.id !== &#39;;&#39; || noreach) {
            return;
        }
        for (;;) {
            t = peek(i);
            if (t.reach) {
                return;
            }
            if (t.id !== &#39;(endline)&#39;) {
                if (t.id === &#39;function&#39;) {
                    warning(
&quot;Inner functions should be listed at the top of the outer function.&quot;,
t);
                    break;
                }
                warning(&quot;Unreachable &#39;&quot; + t.value + &quot;&#39;
after &#39;&quot; + s +
                        &quot;&#39;.&quot;, t);
                break;
            }
            i += 1;
        }
    }


    function statement() {
        var t = token;
        while (t.id === &#39;;&#39;) {
            warning(&quot;Unnecessary semicolon&quot;, t);
            advance(&#39;;&#39;);
            t = token;
            if (t.id === &#39;}&#39;) {
                return;
            }
        }
        if (t.identifier &amp;&amp; !t.reserved &amp;&amp; peek().id ===
&#39;:&#39;) {
            advance();
            advance(&#39;:&#39;);
            addlabel(t.value, &#39;live*&#39;);
            if (!token.labelled) {
                warning(&quot;Label &#39;&quot; + t.value +
                        &quot;&#39; on unlabelable statement &#39;&quot; +
token.value + &quot;&#39;.&quot;,
                        token);
            }
            if (t.value.toLowerCase() === &#39;javascript&#39;) {
                warning(&quot;Label &#39;&quot; + t.value +
                        &quot;&#39; looks like a javascript url.&quot;,
                        token);
            }
            token.label = t.value;
            t = token;
        }
        parse(0, true);
        if (!t.block) {
            if (token.id !== &#39;;&#39;) {
                warning(&quot;Missing &#39;;&#39;&quot;, prevtoken.line,
                        prevtoken.from + prevtoken.value.length);
            } else {
                advance(&#39;;&#39;);
            }
        }
    }


    function statements() {
        while (!token.reach) {
            statement();
        }
    }


    function block() {
        var t = token;
        if (token.id === &#39;{&#39;) {
            advance(&#39;{&#39;);
            statements();
            advance(&#39;}&#39;, t);
        } else {
            warning(&quot;Missing &#39;{&#39; before &#39;&quot; + token.value
+ &quot;&#39;.&quot;);
            noreach = true;
            statement();
            noreach = false;
        }
        verb = null;
    }


// An identity function, used by string and number tokens.

    function idValue() {
        return this;
    }


    function countMember(m) {
        if (typeof member[m] === &#39;number&#39;) {
            member[m] += 1;
        } else {
            member[m] = 1;
        }
    }


// Common HTML attributes that carry scripts.

    var scriptstring = {
        onblur:      true,
        onchange:    true,
        onclick:     true,
        ondblclick:  true,
        onfocus:     true,
        onkeydown:   true,
        onkeypress:  true,
        onkeyup:     true,
        onload:      true,
        onmousedown: true,
        onmousemove: true,
        onmouseout:  true,
        onmouseover: true,
        onmouseup:   true,
        onreset:     true,
        onselect:    true,
        onsubmit:    true,
        onunload:    true
    };


// XML types. Currently we support html and widget.

    var xmltype = {
        HTML: {
            doBegin: function (n) {
                if (!option.cap) {
                    warning(&quot;HTML case error.&quot;);
                }
                xmltype.html.doBegin();
            }
        },
        html: {
            doBegin: function (n) {
                xtype = &#39;html&#39;;
                xmltype.html.script = false;
            },
            doTagName: function (n, p) {
                var i, t = xmltype.html.tag[n], x;
                if (!t) {
                    error(&#39;Unrecognized tag: &lt;&#39; + n + &#39;&gt;.
&#39; +
                            (n === n.toLowerCase() ?
                            &#39;Did you mean &lt;&#39; + n.toLowerCase() +
&#39;&gt;?&#39; : &#39;&#39;));
                }
                x = t.parent;
                if (x) {
                    if (x.indexOf(&#39; &#39; + p + &#39; &#39;) &lt; 0) {
                        error(&#39;A &lt;&#39; + n + &#39;&gt; must be within
&lt;&#39; + x + &#39;&gt;&#39;,
                                prevtoken);
                    }
                } else {
                    i = stack.length;
                    do {
                        if (i &lt;= 0) {
                            error(&#39;A &lt;&#39; + n + &#39;&gt; must be
within the body&#39;,
                                    prevtoken);
                        }
                        i -= 1;
                    } while (stack[i].name !== &#39;body&#39;);
                }
                xmltype.html.script = n === &#39;script&#39;;
                return t.empty;
            },
            doAttribute: function (n, a) {
                if (n === &#39;script&#39;) {
                    if (a === &#39;src&#39;) {
                        xmltype.html.script = false;
                        return &#39;string&#39;;
                    } else if (a === &#39;language&#39;) {
                        warning(&quot;The &#39;language&#39; attribute is
deprecated&quot;,
                                prevtoken);
                        return false;
                    }
                }
                if (a === &#39;href&#39;) {
                    return &#39;href&#39;;
                }
                return scriptstring[a] &amp;&amp; &#39;script&#39;;
            },
            doIt: function (n) {
                return xmltype.html.script ? &#39;script&#39; : n !==
&#39;html&#39; &amp;&amp;
                        xmltype.html.tag[n].special &amp;&amp;
&#39;special&#39;;
            },
            tag: {
                a:        {},
                abbr:     {},
                acronym:  {},
                address:  {},
                applet:   {},
                area:     {empty: true, parent: &#39; map &#39;},
                b:        {},
                base:     {empty: true, parent: &#39; head &#39;},
                bdo:      {},
                big:      {},
                blockquote: {},
                body:     {parent: &#39; html noframes &#39;},
                br:       {empty: true},
                button:   {},
                canvas:   {parent: &#39; body p div th td &#39;},
                caption:  {parent: &#39; table &#39;},
                center:   {},
                cite:     {},
                code:     {},
                col:      {empty: true, parent: &#39; table colgroup &#39;},
                colgroup: {parent: &#39; table &#39;},
                dd:       {parent: &#39; dl &#39;},
                del:      {},
                dfn:      {},
                dir:      {},
                div:      {},
                dl:       {},
                dt:       {parent: &#39; dl &#39;},
                em:       {},
                embed:    {},
                fieldset: {},
                font:     {},
                form:     {},
                frame:    {empty: true, parent: &#39; frameset &#39;},
                frameset: {parent: &#39; html frameset &#39;},
                h1:       {},
                h2:       {},
                h3:       {},
                h4:       {},
                h5:       {},
                h6:       {},
                head:     {parent: &#39; html &#39;},
                html:     {},
                hr:       {empty: true},
                i:        {},
                iframe:   {},
                img:      {empty: true},
                input:    {empty: true},
                ins:      {},
                kbd:      {},
                label:    {},
                legend:   {parent: &#39; fieldset &#39;},
                li:       {parent: &#39; dir menu ol ul &#39;},
                link:     {empty: true, parent: &#39; head &#39;},
                map:      {},
                menu:     {},
                meta:     {empty: true, parent: &#39; head noscript &#39;},
                noframes: {parent: &#39; html body &#39;},
                noscript: {parent: &#39; html head body frameset &#39;},
                object:   {},
                ol:       {},
                optgroup: {parent: &#39; select &#39;},
                option:   {parent: &#39; optgroup select &#39;},
                p:        {},
                param:    {empty: true, parent: &#39; applet object &#39;},
                pre:      {},
                q:        {},
                samp:     {},
                script:   {parent:
&#39; head body p div span abbr acronym address bdo blockquote cite code del
dfn em ins kbd pre samp strong table tbody td th tr var &#39;},
                select:   {},
                small:    {},
                span:     {},
                strong:   {},
                style:    {parent: &#39; head &#39;, special: true},
                sub:      {},
                sup:      {},
                table:    {},
                tbody:    {parent: &#39; table &#39;},
                td:       {parent: &#39; tr &#39;},
                textarea: {},
                tfoot:    {parent: &#39; table &#39;},
                th:       {parent: &#39; tr &#39;},
                thead:    {parent: &#39; table &#39;},
                title:    {parent: &#39; head &#39;},
                tr:       {parent: &#39; table tbody thead tfoot &#39;},
                tt:       {},
                u:        {},
                ul:       {},
                &#39;var&#39;:    {}
            }
        },
        widget: {
            doBegin: function (n) {
                xtype = &#39;widget&#39;;
            },
            doTagName: function (n, p) {
                var t = xmltype.widget.tag[n];
                if (!t) {
                    error(&#39;Unrecognized tag: &lt;&#39; + n + &#39;&gt;.
&#39;);
                }
                var x = t.parent;
                if (x.indexOf(&#39; &#39; + p + &#39; &#39;) &lt; 0) {
                    error(&#39;A &lt;&#39; + n + &#39;&gt; must be within
&lt;&#39; + x + &#39;&gt;&#39;,
                            prevtoken);
                }
            },
            doAttribute: function (n, a) {
                var t = xmltype.widget.tag[a];
                if (!t) {
                    error(&#39;Unrecognized attribute: &lt;&#39; + n + &#39;
&#39; + a + &#39;&gt;. &#39;);
                }
                var x = t.parent;
                if (x.indexOf(&#39; &#39; + n + &#39; &#39;) &lt; 0) {
                    error(&#39;Attribute &#39; + a + &#39; does not belong in
&lt;&#39; +
                            n + &#39;&gt;&#39;);
                }
                return t.script ? &#39;script&#39; : a === &#39;name&#39; ?
&#39;define&#39; : &#39;string&#39;;
            },
            doIt: function (n) {
                var x = xmltype.widget.tag[n];
                return x &amp;&amp; x.script &amp;&amp; &#39;script&#39;;
            },
            tag: {
                &quot;about-box&quot;: {parent: &#39; widget &#39;},
                &quot;about-image&quot;: {parent: &#39; about-box &#39;},
                &quot;about-text&quot;: {parent: &#39; about-box &#39;},
                &quot;about-version&quot;: {parent: &#39; about-box &#39;},
                action: {parent: &#39; widget &#39;, script: true},
                alignment: {parent: &#39; image text textarea window &#39;},
                author: {parent: &#39; widget &#39;},
                autoHide: {parent: &#39; scrollbar &#39;},
                bgColor: {parent: &#39; text textarea &#39;},
                bgOpacity: {parent: &#39; text textarea &#39;},
                checked: {parent: &#39; image menuItem &#39;},
                clipRect: {parent: &#39; image &#39;},
                color: {parent: &#39; about-text about-version shadow text
textarea &#39;},
                contextMenuItems: {parent: &#39; frame image text textarea
window &#39;},
                colorize: {parent: &#39; image &#39;},
                columns: {parent: &#39; textarea &#39;},
                company: {parent: &#39; widget &#39;},
                copyright: {parent: &#39; widget &#39;},
                data: {parent: &#39; about-text about-version text textarea
&#39;},
                debug: {parent: &#39; widget &#39;},
                defaultValue: {parent: &#39; preference &#39;},
                defaultTracking: {parent: &#39; widget &#39;},
                description: {parent: &#39; preference &#39;},
                directory: {parent: &#39; preference &#39;},
                editable: {parent: &#39; textarea &#39;},
                enabled: {parent: &#39; menuItem &#39;},
                extension: {parent: &#39; preference &#39;},
                file: {parent: &#39; action preference &#39;},
                fillMode: {parent: &#39; image &#39;},
                font: {parent: &#39; about-text about-version text textarea
&#39;},
                frame: {parent: &#39; frame window &#39;},
                group: {parent: &#39; preference &#39;},
                hAlign: {parent: &#39; frame image scrollbar text textarea
&#39;},
                height: {parent: &#39; frame image scrollbar text textarea
window &#39;},
                hidden: {parent: &#39; preference &#39;},
                hLineSize: {parent: &#39; frame &#39;},
                hOffset: {parent: &#39; about-text about-version frame image
scrollbar shadow text textarea window &#39;},
                hotkey: {parent: &#39; widget &#39;},
                hRegistrationPoint: {parent: &#39; image &#39;},
                hslAdjustment: {parent: &#39; image &#39;},
                hslTinting: {parent: &#39; image &#39;},
                hScrollBar: {parent: &#39; frame &#39;},
                icon: {parent: &#39; preferenceGroup &#39;},
                id: {parent: &#39; widget &#39;},
                image: {parent: &#39; about-box frame window widget &#39;},
                interval: {parent: &#39; action timer &#39;},
                key: {parent: &#39; hotkey &#39;},
                kind: {parent: &#39; preference &#39;},
                level: {parent: &#39; window &#39;},
                lines: {parent: &#39; textarea &#39;},
                loadingSrc: {parent: &#39; image &#39;},
                max: {parent: &#39; scrollbar &#39;},
                maxLength: {parent: &#39; preference &#39;},
                menuItem: {parent: &#39; contextMenuItems &#39;},
                min: {parent: &#39; scrollbar &#39;},
                minimumVersion: {parent: &#39; widget &#39;},
                minLength: {parent: &#39; preference &#39;},
                missingSrc: {parent: &#39; image &#39;},
                modifier: {parent: &#39; hotkey &#39;},
                name: {parent: &#39; frame hotkey image preference
preferenceGroup scrollbar text textarea timer widget window &#39;},
                notSaved: {parent: &#39; preference &#39;},
                onContextMenu: {parent: &#39; frame image text textarea window
&#39;, script: true},
                onDragDrop: {parent: &#39; frame image text textarea &#39;,
script: true},
                onDragEnter: {parent: &#39; frame image text textarea &#39;,
script: true},
                onDragExit: {parent: &#39; frame image text textarea &#39;,
script: true},
                onFirstDisplay: {parent: &#39; window &#39;, script: true},
                onGainFocus: {parent: &#39; textarea window &#39;, script:
true},
                onKeyDown: {parent: &#39; hotkey text textarea &#39;, script:
true},
                onKeyPress: {parent: &#39; textarea &#39;, script: true},
                onKeyUp: {parent: &#39; hotkey text textarea &#39;, script:
true},
                onImageLoaded: {parent: &#39; image &#39;, script: true},
                onLoseFocus: {parent: &#39; textarea window &#39;, script:
true},
                onMouseDown: {parent: &#39; frame image text textarea &#39;,
script: true},
                onMouseEnter: {parent: &#39; frame image text textarea &#39;,
script: true},
                onMouseExit: {parent: &#39; frame image text textarea &#39;,
script: true},
                onMouseMove: {parent: &#39; frame image text &#39;, script:
true},
                onMouseUp: {parent: &#39; frame image text textarea &#39;,
script: true},
                onMouseWheel: {parent: &#39; frame &#39;, script: true},
                onMultiClick: {parent: &#39; frame image text textarea window
&#39;, script: true},
                onSelect: {parent: &#39; menuItem &#39;, script: true},
                onTimerFired: {parent: &#39; timer &#39;, script: true},
                onValueChanged: {parent: &#39; scrollbar &#39;, script: true},
                opacity: {parent: &#39; frame image scrollbar shadow text
textarea window &#39;},
                option: {parent: &#39; preference widget &#39;},
                optionValue: {parent: &#39; preference &#39;},
                order: {parent: &#39; preferenceGroup &#39;},
                orientation: {parent: &#39; scrollbar &#39;},
                pageSize: {parent: &#39; scrollbar &#39;},
                preference: {parent: &#39; widget &#39;},
                preferenceGroup: {parent: &#39; widget &#39;},
                remoteAsync: {parent: &#39; image &#39;},
                requiredPlatform: {parent: &#39; widget &#39;},
                rotation: {parent: &#39; image &#39;},
                scrollX: {parent: &#39; frame &#39;},
                scrollY: {parent: &#39; frame &#39;},
                secure: {parent: &#39; preference textarea &#39;},
                scrollbar: {parent: &#39; text textarea window &#39;},
                scrolling: {parent: &#39; text &#39;},
                shadow: {parent: &#39; about-text about-version text window
&#39;},
                size: {parent: &#39; about-text about-version text textarea
&#39;},
                spellcheck: {parent: &#39; textarea &#39;},
                src: {parent: &#39; image &#39;},
                srcHeight: {parent: &#39; image &#39;},
                srcWidth: {parent: &#39; image &#39;},
                style: {parent: &#39; about-text about-version preference text
textarea &#39;},
                text: {parent: &#39; frame window &#39;},
                textarea: {parent: &#39; frame window &#39;},
                timer: {parent: &#39; widget &#39;},
                thumbColor: {parent: &#39; scrollbar &#39;},
                ticking: {parent: &#39; timer &#39;},
                ticks: {parent: &#39; preference &#39;},
                tickLabel: {parent: &#39; preference &#39;},
                tileOrigin: {parent: &#39; image &#39;},
                title: {parent: &#39; menuItem preference preferenceGroup
window &#39;},
                tooltip: {parent: &#39; image text textarea &#39;},
                tracking: {parent: &#39; image &#39;},
                trigger: {parent: &#39; action &#39;},
                truncation: {parent: &#39; text &#39;},
                type: {parent: &#39; preference &#39;},
                url: {parent: &#39; about-box about-text about-version &#39;},
                useFileIcon: {parent: &#39; image &#39;},
                vAlign: {parent: &#39; frame image scrollbar text textarea
&#39;},
                value: {parent: &#39; preference scrollbar &#39;},
                version: {parent: &#39; widget &#39;},
                visible: {parent: &#39; frame image scrollbar text textarea
window &#39;},
                vLineSize: {parent: &#39; frame &#39;},
                vOffset: {parent: &#39; about-text about-version frame image
scrollbar shadow text textarea window &#39;},
                vRegistrationPoint: {parent: &#39; image &#39;},
                vScrollBar: {parent: &#39; frame &#39;},
                width: {parent: &#39; frame image scrollbar text textarea
window &#39;},
                window: {parent: &#39; widget &#39;},
                zOrder: {parent: &#39; frame image scrollbar text textarea
&#39;}
            }
        }
    };

    function xmlword(tag) {
        var w = token.value;
        if (!token.identifier) {
            if (token.id === &#39;&lt;&#39;) {
                error(tag ? &quot;Expected &amp;lt; and saw
&#39;&lt;&#39;&quot; : &quot;Missing &#39;&gt;&#39;&quot;,
                        prevtoken);
            } else if (token.id === &#39;(end)&#39;) {
                error(&quot;Bad structure&quot;);
            } else {
                warning(&quot;Missing quotes&quot;, prevtoken);
            }
        }
        advance();
        while (token.id === &#39;-&#39; || token.id === &#39;:&#39;) {
            w += token.id;
            advance();
            if (!token.identifier) {
                error(&#39;Bad name: &#39; + w + token.value);
            }
            w += token.value;
            advance();
        }
        return w;
    }

    function xml() {
        var a, e, n, q, t, v;
        xmode = &#39;xml&#39;;
        stack = [];
        for (;;) {
            switch (token.value) {
            case &#39;&lt;&#39;:
                advance(&#39;&lt;&#39;);
                t = token;
                n = xmlword(true);
                t.name = n;
                if (!xtype) {
                    if (xmltype[n]) {
                        xmltype[n].doBegin();
                        n = xtype;
                        e = false;
                    } else {
                        error(&quot;Unrecognized &lt;&quot; + n +
&quot;&gt;&quot;);
                    }
                } else {
                    if (option.cap &amp;&amp; xtype === &#39;html&#39;) {
                        n = n.toLowerCase();
                    }
                    if (stack.length === 0) {
                        error(&quot;What the hell is this?&quot;, token);
                    }
                    e = xmltype[xtype].doTagName(n,
                            stack[stack.length - 1].type);
                }
                t.type = n;
                for (;;) {
                    if (token.id === &#39;/&#39;) {
                        advance(&#39;/&#39;);
                        e = true;
                        break;
                    }
                    if (token.id &amp;&amp; token.id.substr(0, 1) ===
&#39;&gt;&#39;) {
                        break;
                    }
                    a = xmlword();
                    switch (xmltype[xtype].doAttribute(n, a)) {
                    case &#39;script&#39;:
                        xmode = &#39;string&#39;;
                        advance(&#39;=&#39;);
                        q = token.id;
                        if (q !== &#39;&quot;&#39; &amp;&amp; q !==
&quot;&#39;&quot;) {
                            error(&#39;Missing quote.&#39;);
                        }
                        xmode = q;
                        advance(q);
                        statements();
                        if (token.id !== q) {
                            error(&#39;Missing close quote on script
attribute.&#39;);
                        }
                        xmode = &#39;xml&#39;;
                        advance(q);
                        break;
                    case &#39;value&#39;:
                        advance(&#39;=&#39;);
                        if (!token.identifier &amp;&amp;
                                token.type !== &#39;(string)&#39; &amp;&amp;
                                token.type !== &#39;(number)&#39;) {
                            error(&#39;Bad value: &#39; + token.value);
                        }
                        advance();
                        break;
                    case &#39;string&#39;:
                        advance(&#39;=&#39;);
                        if (token.type !== &#39;(string)&#39;) {
                            error(&#39;Bad value: &#39; + token.value);
                        }
                        advance();
                        break;
                    case &#39;href&#39;:
                        advance(&#39;=&#39;);
                        if (token.type !== &#39;(string)&#39;) {
                            error(&#39;Bad url: &#39; + token.value);
                        }
                        v = token.value.split(&#39;:&#39;);
                        if (v.length &gt; 1) {
                            switch (v[0].substring(0, 4).toLowerCase()) {
                            case &#39;java&#39;:
                            case &#39;jscr&#39;:
                            case &#39;ecma&#39;:
                                warning(&#39;javascript url.&#39;);
                            }
                        }
                        advance();
                        break;
                    case &#39;define&#39;:
                        advance(&#39;=&#39;);
                        if (token.type !== &#39;(string)&#39;) {
                            error(&#39;Bad value: &#39; + token.value);
                        }
                        addlabel(token.value, &#39;global&#39;);
                        advance();
                        break;
                    default:
                        if (token.id === &#39;=&#39;) {
                            advance(&#39;=&#39;);
                            if (!token.identifier &amp;&amp;
                                    token.type !== &#39;(string)&#39;
&amp;&amp;
                                    token.type !== &#39;(number)&#39;) {
                            }
                            advance();
                        }
                    }
                }
                switch (xmltype[xtype].doIt(n)) {
                case &#39;script&#39;:
                    xmode = &#39;script&#39;;
                    advance(&#39;&gt;&#39;);
                    statements();
                    if (token.id !== &#39;&lt;/&#39;) {
                        warning(&quot;Unexpected token.&quot;, token);
                    }
                    xmode = &#39;xml&#39;;
                    break;
                case &#39;special&#39;:
                    e = true;
                    n = &#39;&lt;/&#39; + t.name + &#39;&gt;&#39;;
                    if (!lex.skip(n)) {
                        error(&quot;Missing &quot; + n, t);
                    }
                    break;
                default:
                    lex.skip(&#39;&gt;&#39;);
                }
                if (!e) {
                    stack.push(t);
                }
                break;
            case &#39;&lt;/&#39;:
                advance(&#39;&lt;/&#39;);
                n = xmlword(true);
                t = stack.pop();
                if (!t) {
                    error(&#39;Unexpected close tag: &lt;/&#39; + n +
&#39;&gt;&#39;);
                }
                if (t.name !== n) {
                    error(&#39;Expected &lt;/&#39; + t.name +
                            &#39;&gt; and instead saw &lt;/&#39; + n +
&#39;&gt;&#39;);
                }
                if (token.id !== &#39;&gt;&#39;) {
                    error(&quot;Expected &#39;&gt;&#39;&quot;);
                }
                if (stack.length &gt; 0) {
                    lex.skip(&#39;&gt;&#39;);
                } else {
                    advance(&#39;&gt;&#39;);
                }
                break;
            case &#39;&lt;!&#39;:
                for (;;) {
                    advance();
                    if (token.id === &#39;&gt;&#39;) {
                        break;
                    }
                    if (token.id === &#39;&lt;&#39; || token.id ===
&#39;(end)&#39;) {
                        error(&quot;Missing &#39;&gt;&#39;.&quot;, prevtoken);
                    }
                }
                lex.skip(&#39;&gt;&#39;);
                break;
            case &#39;&lt;!--&#39;:
                lex.skip(&#39;--&gt;&#39;);
                break;
            case &#39;&lt;%&#39;:
                lex.skip(&#39;%&gt;&#39;);
                break;
            case &#39;&lt;?&#39;:
                for (;;) {
                    advance();
                    if (token.id === &#39;?&gt;&#39;) {
                        break;
                    }
                    if (token.id === &#39;&lt;?&#39; || token.id ===
&#39;&lt;&#39; ||
                            token.id === &#39;&gt;&#39; || token.id ===
&#39;(end)&#39;) {
                        error(&quot;Missing &#39;?&gt;&#39;.&quot;, prevtoken);
                    }
                }
                lex.skip(&#39;?&gt;&#39;);
                break;
            case &#39;&lt;=&#39;:
            case &#39;&lt;&lt;&#39;:
            case &#39;&lt;&lt;=&#39;:
                error(&quot;Expected &#39;&amp;lt;&#39;.&quot;);
                break;
            case &#39;(end)&#39;:
                return;
            }
            if (stack.length === 0) {
                return;
            }
            if (!lex.skip(&#39;&#39;)) {
                t = stack.pop();
                error(&#39;Missing &lt;/&#39; + t.name + &#39;&gt;&#39;, t);
            }
            advance();
        }
    }


// Build the syntax table by declaring the syntactic elements of the language.

    type(&#39;(number)&#39;, idValue);
    type(&#39;(string)&#39;, idValue);

    syntax[&#39;(identifier)&#39;] = {
        type: &#39;(identifier)&#39;,
        lbp: 0,
        identifier: true,
        nud: function () {
            if (option.undef &amp;&amp; !builtin(this.value) &amp;&amp;
                    xmode !== &#39;&quot;&#39; &amp;&amp; xmode !==
&quot;&#39;&quot;) {
                var c = funlab;
                while (!c[this.value]) {
                    c = c[&#39;(context)&#39;];
                    if (!c) {
                        warning(&quot;Undefined &quot; +
                                (token.id === &#39;(&#39; ?
&quot;function&quot; : &quot;variable&quot;) +
                                &quot;: &quot; + this.value, prevtoken);
                        break;
                    }
                }
            }
            addlabel(this.value, &#39;global&#39;);
            return this;
        },
        led: function () {
            error(&quot;Expected an operator and instead saw &#39;&quot; +
                token.value + &quot;&#39;.&quot;);
        }
    };

    type(&#39;(regex)&#39;, function () {
        return [this.id, this.value, this.flags];
    });

    delim(&#39;(endline)&#39;);
    delim(&#39;(begin)&#39;);
    delim(&#39;(end)&#39;).reach = true;
    delim(&#39;&lt;/&#39;).reach = true;
    delim(&#39;&lt;![&#39;).reach = true;
    delim(&#39;&lt;%&#39;);
    delim(&#39;&lt;?&#39;);
    delim(&#39;&lt;!&#39;);
    delim(&#39;&lt;!--&#39;);
    delim(&#39;%&gt;&#39;);
    delim(&#39;?&gt;&#39;);
    delim(&#39;(error)&#39;).reach = true;
    delim(&#39;}&#39;).reach = true;
    delim(&#39;)&#39;);
    delim(&#39;]&#39;);
    delim(&#39;]]&gt;&#39;).reach = true;
    delim(&#39;&quot;&#39;).reach = true;
    delim(&quot;&#39;&quot;).reach = true;
    delim(&#39;;&#39;);
    delim(&#39;:&#39;).reach = true;
    delim(&#39;,&#39;);
    reservevar(&#39;eval&#39;);
    reserve(&#39;else&#39;);
    reserve(&#39;case&#39;).reach = true;
    reserve(&#39;default&#39;).reach = true;
    reserve(&#39;catch&#39;);
    reserve(&#39;finally&#39;);
    reservevar(&#39;arguments&#39;);
    reservevar(&#39;false&#39;);
    reservevar(&#39;Infinity&#39;);
    reservevar(&#39;NaN&#39;);
    reservevar(&#39;null&#39;);
    reservevar(&#39;this&#39;);
    reservevar(&#39;true&#39;);
    reservevar(&#39;undefined&#39;);
    assignop(&#39;=&#39;, &#39;assign&#39;, 20);
    assignop(&#39;+=&#39;, &#39;assignadd&#39;, 20);
    assignop(&#39;-=&#39;, &#39;assignsub&#39;, 20);
    assignop(&#39;*=&#39;, &#39;assignmult&#39;, 20);
    assignop(&#39;/=&#39;, &#39;assigndiv&#39;, 20).nud = function () {
        warning(
                &quot;A regular expression literal can be confused with
&#39;/=&#39;.&quot;);
    };
    assignop(&#39;%=&#39;, &#39;assignmod&#39;, 20);
    assignop(&#39;&amp;=&#39;, &#39;assignbitand&#39;, 20);
    assignop(&#39;|=&#39;, &#39;assignbitor&#39;, 20);
    assignop(&#39;^=&#39;, &#39;assignbitxor&#39;, 20);
    assignop(&#39;&lt;&lt;=&#39;, &#39;assignshiftleft&#39;, 20);
    assignop(&#39;&gt;&gt;=&#39;, &#39;assignshiftright&#39;, 20);
    assignop(&#39;&gt;&gt;&gt;=&#39;, &#39;assignshiftrightunsigned&#39;, 20);
    infix(&#39;?&#39;, function (left) {
        parse(10);
        advance(&#39;:&#39;);
        parse(10);
    }, 30);

    infix(&#39;||&#39;, &#39;or&#39;, 40);
    infix(&#39;&amp;&amp;&#39;, &#39;and&#39;, 50);
    infix(&#39;|&#39;, &#39;bitor&#39;, 70);
    infix(&#39;^&#39;, &#39;bitxor&#39;, 80);
    infix(&#39;&amp;&#39;, &#39;bitand&#39;, 90);
    infix(&#39;==&#39;, function (left) {
        var t = token;
        if (option.eqeqeq) {
            warning(&quot;Use &#39;===&#39; instead of &#39;==&#39;.&quot;, t);
        } else if (    (t.type === &#39;(number)&#39; &amp;&amp; !+t.value) ||
                (t.type === &#39;(string)&#39; &amp;&amp; !t.value) ||
                t.type === &#39;true&#39; || t.type === &#39;false&#39; ||
                t.type === &#39;undefined&#39; || t.type === &#39;null&#39;) {
            warning(&quot;Use &#39;===&#39; to compare with &#39;&quot; +
t.value + &quot;&#39;.&quot;, t);
        }
        return [&#39;==&#39;, left, parse(100)];
    }, 100);
    infix(&#39;===&#39;, &#39;equalexact&#39;, 100);
    infix(&#39;!=&#39;, function (left) {
        var t = token;
        if (option.eqeqeq) {
            warning(&quot;Use &#39;!==&#39; instead of &#39;!=&#39;.&quot;, t);
        } else if (    (t.type === &#39;(number)&#39; &amp;&amp; !+t.value) ||
                (t.type === &#39;(string)&#39; &amp;&amp; !t.value) ||
                t.type === &#39;true&#39; || t.type === &#39;false&#39; ||
                t.type === &#39;undefined&#39; || t.type === &#39;null&#39;) {
            warning(&quot;Use &#39;!==&#39; to compare with &#39;&quot; +
t.value + &quot;&#39;.&quot;, t);
        }
        return [&#39;!=&#39;, left, parse(100)];
    }, 100);
    infix(&#39;!==&#39;, &#39;notequalexact&#39;, 100);
    infix(&#39;&lt;&#39;, &#39;less&#39;, 110);
    infix(&#39;&gt;&#39;, &#39;greater&#39;, 110);
    infix(&#39;&lt;=&#39;, &#39;lessequal&#39;, 110);
    infix(&#39;&gt;=&#39;, &#39;greaterequal&#39;, 110);
    infix(&#39;&lt;&lt;&#39;, &#39;shiftleft&#39;, 120);
    infix(&#39;&gt;&gt;&#39;, &#39;shiftright&#39;, 120);
    infix(&#39;&gt;&gt;&gt;&#39;, &#39;shiftrightunsigned&#39;, 120);
    infix(&#39;in&#39;, &#39;in&#39;, 120);
    infix(&#39;instanceof&#39;, &#39;instanceof&#39;, 120);
    infix(&#39;+&#39;, &#39;addconcat&#39;, 130);
    prefix(&#39;+&#39;, &#39;num&#39;);
    infix(&#39;-&#39;, &#39;sub&#39;, 130);
    prefix(&#39;-&#39;, &#39;neg&#39;);
    infix(&#39;*&#39;, &#39;mult&#39;, 140);
    infix(&#39;/&#39;, &#39;div&#39;, 140);
    infix(&#39;%&#39;, &#39;mod&#39;, 140);

    suffix(&#39;++&#39;, &#39;postinc&#39;);
    prefix(&#39;++&#39;, &#39;preinc&#39;);
    syntax[&#39;++&#39;].exps = true;

    suffix(&#39;--&#39;, &#39;postdec&#39;);
    prefix(&#39;--&#39;, &#39;predec&#39;);
    syntax[&#39;--&#39;].exps = true;
    prefixname(&#39;delete&#39;, function () {
        parse(0);
    }).exps = true;


    prefix(&#39;~&#39;, &#39;bitnot&#39;);
    prefix(&#39;!&#39;, &#39;not&#39;);
    prefixname(&#39;typeof&#39;, &#39;typeof&#39;);
    prefixname(&#39;new&#39;, function () {
        var c = parse(155),
            i;
        if (c) {
            if (c.identifier) {
                c[&#39;new&#39;] = true;
                switch (c.value) {
                case &#39;Object&#39;:
                    warning(&#39;Use the object literal notation {}.&#39;,
prevtoken);
                    break;
                case &#39;Array&#39;:
                    warning(&#39;Use the array literal notation [].&#39;,
prevtoken);
                    break;
                case &#39;Number&#39;:
                case &#39;String&#39;:
                case &#39;Boolean&#39;:
                    warning(&quot;Do not use the &quot; + c.value +
                        &quot; function as a constructor.&quot;, prevtoken);
                    break;
                case &#39;Function&#39;:
                    if (!option.evil) {
                        warning(&#39;The Function constructor is eval.&#39;);
                    }
                    break;
                default:
                    i = c.value.substr(0, 1);
                    if (i &lt; &#39;A&#39; || i &gt; &#39;Z&#39;) {
                        warning(
                &#39;A constructor name should start with an uppercase
letter.&#39;, c);
                    }
                }
            } else {
                if (c.id !== &#39;.&#39; &amp;&amp; c.id !== &#39;[&#39;
&amp;&amp; c.id !== &#39;(&#39;) {
                    warning(&#39;Bad constructor&#39;, prevtoken);
                }
            }
        } else {
            warning(&quot;Weird construction. Delete &#39;new&#39;.&quot;,
this);
        }
        if (token.id === &#39;(&#39;) {
            advance(&#39;(&#39;);
            if (token.id !== &#39;)&#39;) {
                for (;;) {
                    parse(10);
                    if (token.id !== &#39;,&#39;) {
                        break;
                    }
                    advance(&#39;,&#39;);
                }
            }
            advance(&#39;)&#39;);
        } else {
            warning(&quot;Missing &#39;()&#39; invoking a constructor.&quot;);
        }
        return syntax[&#39;function&#39;];
    });
    syntax[&#39;new&#39;].exps = true;

    infix(&#39;.&#39;, function (left) {
        var m = identifier();
        if (typeof m === &#39;string&#39;) {
            countMember(m);
        }
        if (!option.evil &amp;&amp; left &amp;&amp; left.value ===
&#39;document&#39; &amp;&amp;
                (m === &#39;write&#39; || m === &#39;writeln&#39;)) {
            warning(&quot;document.write can be a form of eval.&quot;, left);
        }
        this.left = left;
        this.right = m;
        return this;
    }, 160);

    infix(&#39;(&#39;, function (left) {
        var n = 0, p = [];
        if (left &amp;&amp; left.type === &#39;(identifier)&#39;) {
            if (left.value.match(/^[A-Z](.*[a-z].*)?$/)) {
                if (left.value !== &#39;Number&#39; &amp;&amp; left.value !==
&#39;String&#39;) {
                    warning(&quot;Missing &#39;new&#39; prefix when invoking a
constructor&quot;,
                            left);
                }
            }
        }
        if (token.id !== &#39;)&#39;) {
            for (;;) {
                p[p.length] = parse(10);
                n += 1;
                if (token.id !== &#39;,&#39;) {
                    break;
                }
                advance(&#39;,&#39;);
            }
        }
        advance(&#39;)&#39;);
        if (typeof left === &#39;object&#39;) {
            if (left.value === &#39;parseInt&#39; &amp;&amp; n === 1) {
                warning(&quot;Missing radix parameter&quot;, left);
            }
            if (!option.evil) {
                if (left.value === &#39;eval&#39; || left.value ===
&#39;Function&#39;) {
                    if (    p[0][0] !== &quot;addconcat&quot; ||
                            p[0][2].value !== &#39;)&#39; ||
                            p[0][1][0] !== &quot;addconcat&quot; ||
                            p[0][1][1].value !== &#39;(&#39;) {
                        warning(&quot;eval is evil&quot;, left);
                    }
                } else if (p[0] &amp;&amp; p[0].id === &#39;(string)&#39;
&amp;&amp;
                       (left.value === &#39;setTimeout&#39; ||
                        left.value === &#39;setInterval&#39;)) {
                    warning(
    &quot;Implied eval is evil. Use a function argument instead of a
string&quot;, left);
                }
            }
            if (!left.identifier &amp;&amp; left.id !== &#39;.&#39; &amp;&amp;
                    left.id !== &#39;[&#39; &amp;&amp; left.id !== &#39;(&#39;)
{
                warning(&#39;Bad invocation.&#39;, left);
            }

        }
        return syntax[&#39;function&#39;];
    }, 155).exps = true;

    prefix(&#39;(&#39;, function () {
        parse(0);
        advance(&#39;)&#39;, this);
    });

    infix(&#39;[&#39;, function (left) {
        var e = parse(0);
        if (e &amp;&amp; e.type === &#39;(string)&#39;) {
            countMember(e.value);
            if (ix.test(e.value)) {
                var s = syntax[e.value];
                if (!s || !s.reserved) {
                    warning(&quot;This is better written in dot
notation.&quot;, e);
                }
            }
        }
        advance(&#39;]&#39;, this);
        this.left = left;
        this.right = e;
        return this;
    }, 160);

    prefix(&#39;[&#39;, function () {
        if (token.id === &#39;]&#39;) {
            advance(&#39;]&#39;);
            return;
        }
        for (;;) {
            parse(10);
            if (token.id === &#39;,&#39;) {
                advance(&#39;,&#39;);
                if (token.id === &#39;]&#39; || token.id === &#39;,&#39;) {
                    warning(&#39;Extra comma.&#39;, prevtoken);
                }
            } else {
                advance(&#39;]&#39;, this);
                return;
            }
        }
    }, 160);

    (function (x) {
        x.nud = function () {
            var i;
            if (token.id === &#39;}&#39;) {
                advance(&#39;}&#39;);
                return;
            }
            for (;;) {
                i = optionalidentifier(true);
                if (!i &amp;&amp; (token.id === &#39;(string)&#39; ||
                           token.id === &#39;(number)&#39;)) {
                    i = token.id;
                    advance();
                }
                if (!i) {
                    error(&quot;Expected an identifier or &#39;}&#39; and
instead saw &#39;&quot; +
                            token.value + &quot;&#39;.&quot;);
                }
                if (typeof i.value === &#39;string&#39;) {
                    countMember(i.value);
                }
                advance(&#39;:&#39;);
                parse(10);
                if (token.id === &#39;,&#39;) {
                    advance(&#39;,&#39;);
                    if (token.id === &#39;,&#39; || token.id === &#39;}&#39;) {
                        warning(&quot;Extra comma.&quot;);
                    }
                } else {
                    advance(&#39;}&#39;, this);
                    return;
                }
            }
        };
        x.fud = function () {
            error(&quot;Expected to see a statement and instead saw a
block.&quot;);
        };
    })(delim(&#39;{&#39;));


    function varstatement() {
        var i, n;
        for (;;) {
            n = identifier();
            if (!option.redef) {
                for (i = funstack.length - 1; i &gt;= 0; i -= 1) {
                    if (funstack[i][n]) {
                        warning(&quot;Redefinition of &#39;&quot; + n +
&quot;&#39;.&quot;, prevtoken);
                        break;
                    }
                }
            }
            addlabel(n, &#39;var&#39;);
            if (token.id === &#39;=&#39;) {
                advance(&#39;=&#39;);
                parse(20);
            }
            if (token.id === &#39;,&#39;) {
                advance(&#39;,&#39;);
            } else {
                return;
            }
        }
    }


    stmt(&#39;var&#39;, varstatement);

    stmt(&#39;new&#39;, function () {
        error(&quot;&#39;new&#39; should not be used as a statement.&quot;);
    });


    function functionparams() {
        var t = token;
        advance(&#39;(&#39;);
        if (token.id === &#39;)&#39;) {
            advance(&#39;)&#39;);
            return;
        }
        for (;;) {
            addlabel(identifier(), &#39;parameter&#39;);
            if (token.id === &#39;,&#39;) {
                advance(&#39;,&#39;);
            } else {
                advance(&#39;)&#39;, t);
                return;
            }
        }
    }


    blockstmt(&#39;function&#39;, function () {
        var i = identifier();
        addlabel(i, &#39;var*&#39;);
        beginfunction(i);
        addlabel(i, &#39;function&#39;);
        functionparams();
        block();
        endfunction();
        if (token.id === &#39;(&#39; &amp;&amp; token.line === prevtoken.line)
{
            error(
&quot;Function statements are not invocable. Wrap the function expression in
parens.&quot;);
        }
    });

    prefixname(&#39;function&#39;, function () {
        var i = optionalidentifier() || (&#39;&quot;&#39; + anonname +
&#39;&quot;&#39;);
        beginfunction(i);
        addlabel(i, &#39;function&#39;);
        functionparams();
        block();
        endfunction();
    });

    blockstmt(&#39;if&#39;, function () {
        var t = token;
        advance(&#39;(&#39;);
        parse(20);
        if (token.id === &#39;=&#39;) {
            warning(&quot;Assignment in control part.&quot;);
            advance(&#39;=&#39;);
            parse(20);
        }
        advance(&#39;)&#39;, t);
        block();
        if (token.id === &#39;else&#39;) {
            advance(&#39;else&#39;);
            if (token.id === &#39;if&#39; || token.id === &#39;switch&#39;) {
                statement();
            } else {
                block();
            }
        }
    });

    blockstmt(&#39;try&#39;, function () {
        var b;
        block();
        if (token.id === &#39;catch&#39;) {
            advance(&#39;catch&#39;);
            beginfunction(&#39;&quot;catch&quot;&#39;);
            functionparams();
            block();
            endfunction();
            b = true;
        }
        if (token.id === &#39;finally&#39;) {
            advance(&#39;finally&#39;);
            beginfunction(&#39;&quot;finally&quot;&#39;);
            block();
            endfunction();
            return;
        } else if (!b) {
            error(&quot;Expected &#39;catch&#39; or &#39;finally&#39; and
instead saw &#39;&quot; +
                    token.value + &quot;&#39;.&quot;);
        }
    });

    blockstmt(&#39;while&#39;, function () {
        var t= token;
        advance(&#39;(&#39;);
        parse(20);
        if (token.id === &#39;=&#39;) {
            warning(&quot;Assignment in control part.&quot;);
            advance(&#39;=&#39;);
            parse(20);
        }
        advance(&#39;)&#39;, t);
        block();
    }).labelled = true;

    reserve(&#39;with&#39;);

    blockstmt(&#39;switch&#39;, function () {
        var t = token, g = false;
        advance(&#39;(&#39;);
        this.condition = parse(20);
        advance(&#39;)&#39;, t);
        t = token;
        advance(&#39;{&#39;);
        this.cases = [];
        for (;;) {
            switch (token.id) {
            case &#39;case&#39;:
                switch (verb) {
                case &#39;break&#39;:
                case &#39;case&#39;:
                case &#39;continue&#39;:
                case &#39;return&#39;:
                case &#39;switch&#39;:
                case &#39;throw&#39;:
                    break;
                default:
                    warning(
                        &quot;Expected a &#39;break&#39; statement before
&#39;case&#39;.&quot;,
                        prevtoken);
                }
                advance(&#39;case&#39;);
                this.cases.push(parse(20));
                g = true;
                advance(&#39;:&#39;);
                verb = &#39;case&#39;;
                break;
            case &#39;default&#39;:
                switch (verb) {
                case &#39;break&#39;:
                case &#39;continue&#39;:
                case &#39;return&#39;:
                case &#39;throw&#39;:
                    break;
                default:
                    warning(
                        &quot;Expected a &#39;break&#39; statement before
&#39;default&#39;.&quot;,
                        prevtoken);
                }
                advance(&#39;default&#39;);
                g = true;
                advance(&#39;:&#39;);
                break;
            case &#39;}&#39;:
                advance(&#39;}&#39;, t);
                if (this.cases.length === 1 || this.condition.id ===
&#39;true&#39; ||
                        this.condition.id === &#39;false&#39;) {
                    warning(&quot;The switch should be an if&quot;, this);
                }
                return;
            case &#39;(end)&#39;:
                error(&quot;Missing &#39;}&#39;.&quot;);
                return;
            default:
                if (g) {
                    switch (prevtoken.id) {
                    case &#39;,&#39;:
                        error(&quot;Each value should have its own case
label.&quot;);
                        return;
                    case &#39;:&#39;:
                        statements();
                        break;
                    default:
                        error(&quot;Missing &#39;:&#39; on a case
clause.&quot;, prevtoken);
                    }
                } else {
                    error(&quot;Expected to see &#39;case&#39; and instead saw
&#39;&quot; +
                        token.value + &quot;&#39;.&quot;);
                }
            }
        }
    }).labelled = true;

    stmt(&#39;debugger&#39;, function () {
        if (!option.debug) {
            warning(&quot;All debugger statements should be removed.&quot;);
        }
    });

    stmt(&#39;do&#39;, function () {
        block();
        advance(&#39;while&#39;);
        var t = token;
        advance(&#39;(&#39;);
        parse(20);
        advance(&#39;)&#39;, t);
    }).labelled = true;

    blockstmt(&#39;for&#39;, function () {
        var t = token;
        advance(&#39;(&#39;);
        if (peek(token.id === &#39;var&#39; ? 1 : 0).id === &#39;in&#39;) {
            if (token.id === &#39;var&#39;) {
                advance(&#39;var&#39;);
                addlabel(identifier(), &#39;var&#39;);
            } else {
                advance();
            }
            advance(&#39;in&#39;);
            parse(20);
            advance(&#39;)&#39;, t);
            block();
            return;
        } else {
            if (token.id !== &#39;;&#39;) {
                if (token.id === &#39;var&#39;) {
                    advance(&#39;var&#39;);
                    varstatement();
                } else {
                    for (;;) {
                        parse(0);
                        if (token.id !== &#39;,&#39;) {
                            break;
                        }
                        advance(&#39;,&#39;);
                    }
                }
            }
            advance(&#39;;&#39;);
            if (token.id !== &#39;;&#39;) {
                parse(20);
            }
            advance(&#39;;&#39;);
            if (token.id === &#39;;&#39;) {
                error(&quot;Expected to see &#39;)&#39; and instead saw
&#39;;&#39;.&quot;);
            }
            if (token.id !== &#39;)&#39;) {
                for (;;) {
                    parse(0);
                    if (token.id !== &#39;,&#39;) {
                        break;
                    }
                    advance(&#39;,&#39;);
                }
            }
            advance(&#39;)&#39;, t);
            block();
        }
    }).labelled = true;


    function nolinebreak(t) {
        if (t.line !== token.line) {
            warning(&quot;Statement broken badly.&quot;, t);
        }
    }


    stmt(&#39;break&#39;, function () {
        nolinebreak(this);
        if (funlab[token.value] === &#39;live*&#39;) {
            advance();
        }
        reachable(&#39;break&#39;);
    });


    stmt(&#39;continue&#39;, function () {
        nolinebreak(this);
        if (funlab[token.id] === &#39;live*&#39;) {
            advance();
        }
        reachable(&#39;continue&#39;);
    });


    stmt(&#39;return&#39;, function () {
        nolinebreak(this);
        if (token.id !== &#39;;&#39; &amp;&amp; !token.reach) {
            parse(20);
        }
        reachable(&#39;return&#39;);
    });


    stmt(&#39;throw&#39;, function () {
        nolinebreak(this);
        parse(20);
        reachable(&#39;throw&#39;);
    });


//  Superfluous reserved words

    reserve(&#39;abstract&#39;);
    reserve(&#39;boolean&#39;);
    reserve(&#39;byte&#39;);
    reserve(&#39;char&#39;);
    reserve(&#39;class&#39;);
    reserve(&#39;const&#39;);
    reserve(&#39;double&#39;);
    reserve(&#39;enum&#39;);
    reserve(&#39;export&#39;);
    reserve(&#39;extends&#39;);
    reserve(&#39;final&#39;);
    reserve(&#39;float&#39;);
    reserve(&#39;goto&#39;);
    reserve(&#39;implements&#39;);
    reserve(&#39;import&#39;);
    reserve(&#39;int&#39;);
    reserve(&#39;interface&#39;);
    reserve(&#39;long&#39;);
    reserve(&#39;native&#39;);
    reserve(&#39;package&#39;);
    reserve(&#39;private&#39;);
    reserve(&#39;protected&#39;);
    reserve(&#39;public&#39;);
    reserve(&#39;short&#39;);
    reserve(&#39;static&#39;);
    reserve(&#39;super&#39;);
    reserve(&#39;synchronized&#39;);
    reserve(&#39;throws&#39;);
    reserve(&#39;transient&#39;);
    reserve(&#39;void&#39;);
    reserve(&#39;volatile&#39;);


// The actual JSLINT function itself.

    var itself = function (s, o) {
        option = o;
        if (!o) {
            option = {};
        }
        JSLINT.errors = [];
        globals = {};
        functions = [];
        xmode = false;
        xtype = &#39;&#39;;
        stack = null;
        funlab = {};
        member = {};
        funstack = [];
        lookahead = [];
        lex.init(s);

        prevtoken = token = syntax[&#39;(begin)&#39;];
        try {
            advance();
            if (token.value.charAt(0) === &#39;&lt;&#39;) {
                xml();
            } else {
                statements();
            }
            advance(&#39;(end)&#39;);
        } catch (e) {
            if (e) {
                JSLINT.errors.push({
                    reason: &quot;JSLint error: &quot; + e.description,
                    line: token.line,
                    character: token.from,
                    evidence: token.value
                });
            }
        }
        return JSLINT.errors.length === 0;
    };


// Report generator.

    itself.report = function (option) {
        var a = [], c, cc, f, i, k, o = [], s;

        function detail(h) {
            if (s.length) {
                o.push(&#39;&lt;div&gt;&#39; + h + &#39;:&amp;nbsp; &#39; +
s.sort().join(&#39;, &#39;) +
                    &#39;&lt;/div&gt;&#39;);
            }
        }

        k = JSLINT.errors.length;
        if (k) {
            //o.push(
            //    &#39;&lt;div id=errors&gt;Error:&lt;blockquote&gt;&#39;);
            for (i = 0; i &lt; k; i += 1) {
                c = JSLINT.errors[i];
                if (c) {   
                      o.push(""&lt;p&gt;"" + (c.line) + &#39;|&#39; +
(c.character + 1) + "|" + c.reason.entityify() + "|" +
c.evidence.entityify() + "";&lt;/p&gt;"");  
                }
            }
            //o.push(&#39;&lt;/blockquote&gt;&lt;/div&gt;&#39;);
            if (!c) {
                return o.join(&#39;&#39;);
            }
        }

        if (!option) {
            for (k in member) {
                a.push(k);
            }
            if (a.length) {
                a = a.sort();
                o.push(
                
&#39;&lt;table&gt;&lt;tbody&gt;&lt;tr&gt;&lt;th&gt;Members&lt;/th&gt;&lt;th&gt;Occurrences&lt;/th&gt;&lt;/tr&gt;&#39;);
                for (i = 0; i &lt; a.length; i += 1) {
                    o.push(&#39;&lt;tr&gt;&lt;td&gt;&lt;tt&gt;&#39;, a[i],
&#39;&lt;/tt&gt;&lt;/td&gt;&lt;td&gt;&#39;, member[a[i]],
                            &#39;&lt;/td&gt;&lt;/tr&gt;&#39;);
                }
                o.push(&#39;&lt;/tbody&gt;&lt;/table&gt;&#39;);
            }
            for (i = 0; i &lt; functions.length; i += 1) {
                f = functions[i];
                for (k in f) {
                    if (f[k] === &#39;global&#39;) {
                        c = f[&#39;(context)&#39;];
                        for (;;) {
                            cc = c[&#39;(context)&#39;];
                            if (!cc) {
                                if ((!funlab[k] || funlab[k] ===
&#39;var?&#39;) &amp;&amp;
                                        !builtin(k)) {
                                    funlab[k] = &#39;var?&#39;;
                                    f[k] = &#39;global?&#39;;
                                }
                                break;
                            }
                            if (c[k] === &#39;parameter!&#39; || c[k] ===
&#39;var!&#39;) {
                                f[k] = &#39;var.&#39;;
                                break;
                            }
                            if (c[k] === &#39;var&#39; || c[k] ===
&#39;var*&#39; ||
                                    c[k] === &#39;var!&#39;) {
                                f[k] = &#39;var.&#39;;
                                c[k] = &#39;var!&#39;;
                                break;
                            }
                            if (c[k] === &#39;parameter&#39;) {
                                f[k] = &#39;var.&#39;;
                                c[k] = &#39;parameter!&#39;;
                                break;
                            }
                            c = cc;
                        }
                    }
                }
            }
            s = [];
            for (k in funlab) {
                c = funlab[k];
                if (typeof c === &#39;string&#39; &amp;&amp; c.substr(0, 3) ===
&#39;var&#39;) {
                    if (c === &#39;var?&#39;) {
                        s.push(&#39;&lt;tt&gt;&#39; + k +
&#39;&lt;/tt&gt;&lt;small&gt;&amp;nbsp;(?)&lt;/small&gt;&#39;);
                    } else {
                        s.push(&#39;&lt;tt&gt;&#39; + k +
&#39;&lt;/tt&gt;&#39;);
                    }
                }
            }
            detail(&#39;Global&#39;);
            if (functions.length) {
                o.push(&#39;&lt;br&gt;Function:&lt;ol
style=&quot;padding-left:0.5in&quot;&gt;&#39;);
            }
            for (i = 0; i &lt; functions.length; i += 1) {
                f = functions[i];
                o.push(&#39;&lt;li value=&#39; +
                        f[&#39;(line)&#39;] + &#39;&gt;&lt;tt&gt;&#39; +
(f[&#39;(name)&#39;] || &#39;&#39;) + &#39;&lt;/tt&gt;&#39;);
                s = [];
                for (k in f) {
                    if (k.charAt(0) !== &#39;(&#39;) {
                        switch (f[k]) {
                        case &#39;parameter&#39;:
                            s.push(&#39;&lt;tt&gt;&#39; + k +
&#39;&lt;/tt&gt;&#39;);
                            break;
                        case &#39;parameter!&#39;:
                            s.push(&#39;&lt;tt&gt;&#39; + k +
                                   
&#39;&lt;/tt&gt;&lt;small&gt;&amp;nbsp;(closure)&lt;/small&gt;&#39;);
                            break;
                        }
                    }
                }
                detail(&#39;Parameter&#39;);
                s = [];
                for (k in f) {
                    if (k.charAt(0) !== &#39;(&#39;) {
                        switch(f[k]) {
                        case &#39;var&#39;:
                            s.push(&#39;&lt;tt&gt;&#39; + k +
                                   
&#39;&lt;/tt&gt;&lt;small&gt;&amp;nbsp;(unused)&lt;/small&gt;&#39;);
                            break;
                        case &#39;var*&#39;:
                            s.push(&#39;&lt;tt&gt;&#39; + k +
&#39;&lt;/tt&gt;&#39;);
                            break;
                        case &#39;var!&#39;:
                            s.push(&#39;&lt;tt&gt;&#39; + k +
                                   
&#39;&lt;/tt&gt;&lt;small&gt;&amp;nbsp;(closure)&lt;/small&gt;&#39;);
                            break;
                        case &#39;var.&#39;:
                            s.push(&#39;&lt;tt&gt;&#39; + k +
                                   
&#39;&lt;/tt&gt;&lt;small&gt;&amp;nbsp;(outer)&lt;/small&gt;&#39;);
                            break;
                        }
                    }
                }
                detail(&#39;Var&#39;);
                s = [];
                c = f[&#39;(context)&#39;];
                for (k in f) {
                    if (k.charAt(0) !== &#39;(&#39; &amp;&amp; f[k].substr(0,
6) === &#39;global&#39;) {
                        if (f[k] === &#39;global?&#39;) {
                            s.push(&#39;&lt;tt&gt;&#39; + k +
                                   
&#39;&lt;/tt&gt;&lt;small&gt;&amp;nbsp;(?)&lt;/small&gt;&#39;);
                        } else {
                            s.push(&#39;&lt;tt&gt;&#39; + k +
&#39;&lt;/tt&gt;&#39;);
                        }
                    }
                }
                detail(&#39;Global&#39;);
                s = [];
                for (k in f) {
                    if (k.charAt(0) !== &#39;(&#39; &amp;&amp; f[k] ===
&#39;label&#39;) {
                       s.push(k);
                    }
                }
                detail(&#39;Label&#39;);
                o.push(&#39;&lt;/li&gt;&#39;);
            }
            if (functions.length) {
                o.push(&#39;&lt;/ol&gt;&#39;);
            }
        }
        return o.join(&#39;&#39;);
    };

    return itself;

}();
</pre>
</body>
</html>