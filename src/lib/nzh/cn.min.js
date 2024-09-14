/*
 * @Author: Marlon
 * @Date: 2024-09-14 13:44:46
 * @Description: 
 */
var getNzhObjByLang = require("./src/autoGet");
var langs = {
	s: require("./src/langs/cn_s"),
	b: require("./src/langs/cn_b"),
}
module.exports = getNzhObjByLang(langs.s, langs.b); 