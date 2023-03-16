/*
	我爱模板网 www.5imoban.net
*/

// 语言设置
const languageMap = {
	ch: {
		description: 'CN',
		hello: '你好',
		world: '世界',
		fill:'填写问题',
		sysname:'888报修系统',
		statusfix:'状态维护',
		status:'状态',
		chargefix:'负责人维护',
	},
	en: {
		description: 'EN',
		hello: 'Hello',
		world: 'World',
		fill:'Fill In the problem',
		sysname:'888 repair system',
		statusfix:'Status maintenance',
		status:'status',
		chargefix:'Responsible for maintenance'
	}
}

function changeLanguage(language){
    localStorage.setItem('language', language);
    for(let key in languageMap){
		if(key === language){
    		document.querySelector('#languageStatus').innerText = languageMap[key].description;
    		break;
    	}
	}
	setLanguage();
}

function setLanguage(){
	let elements = document.querySelectorAll('[language]');
	for(let i=0; i<elements.length; i++){
		let curLanVal = elements[i].getAttribute('language');
		for(let key in languageMap){
			let curLanType = localStorage.getItem('language') || 'ch';
			if(key === curLanType){
	    		elements[i].innerText = languageMap[key][curLanVal];
	    		break;
	    	}
		}
	}
}

function getLanguage(param){
	for(let key in languageMap){
		let curLanType = localStorage.getItem('language') || 'ch';
		if(key === curLanType){
			return languageMap[key][param];
		}
	}
}