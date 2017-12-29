RELS.IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
RELS.DRAW = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
/* 20.5 DrawingML - SpreadsheetML Drawing */
function parse_drawing(data, rels/*:any*/) {
	if(!data) return "??";
	/*
		Chartsheet Drawing:
		 - 20.5.2.35 wsDr CT_Drawing
			- 20.5.2.1  absoluteAnchor CT_AbsoluteAnchor
			 - 20.5.2.16 graphicFrame CT_GraphicalObjectFrame
				- 20.1.2.2.16 graphic CT_GraphicalObject
				 - 20.1.2.2.17 graphicData CT_GraphicalObjectData
					- chart reference
		 the actual type is based on the URI of the graphicData
		TODO: handle embedded charts and other types of graphics
	*/
	var id = (data.match(/<c:chart [^>]*r:id="([^"]*)"/)||["",""])[1];

	return rels['!id'][id].Target;
}

/* FEATURE[russa]: parse drawing xml */
/* 20.5 DrawingML - SpreadsheetML Drawing */
function parse_drawing_images(zip, path, data, refs, images, opts, themes, styles) {
	var anchors = [];
	var stack = [];
	var anchor, pic, pos, img;
	var isIgnore = false;

	var pushPos = function(match, type){
		var i = match.index + match[0].length;
		var pos = {};
		pos[type] = i;
		stack.push(pos);
	}

	var parseNum = function(curMatch, type){
		var lastPos = stack.pop();
		if(typeof lastPos[type] === 'number'){
			var numStr = curMatch.input.substring(lastPos[type], curMatch.index).trim();
			return parseInt(numStr, 10);
		}// else if(opts)...
	}

	var errorString = function(tag, type, match, anchorOrPic, pos){
		var errStr = 'Error parsing '+tag+' (in '+path+' at ['+match.index+']):';
		if(type === 'anchor'){
			return errStr+(anchorOrPic? '' : ' <anchor missing>');
		} else if(type === 'position'){
			return errStr+(pos? '' : ' <position information missing>');
		} else if(type === 'both'){
			return errStr+(anchorOrPic? '' : ' <anchor missing>')+(pos? '' : ' <position information missing>');
		} else if(type === 'picture'){
			return errStr+(anchorOrPic? '' : ' <picture element missing>');
		} else if(type === 'image'){
			return errStr+(anchorOrPic? '' : ' <could not parse image data>');
		}
		return errStr;
	}

	var match, x, lastPos;
	tagregex.lastIndex = 0;
	while(match = tagregex.exec(data)){
		x = match[0];
		var y = parsexmltag(x);
		var tag = y[0];
		var len = tag.length;
		tag = tag[len-1] === '>' && tag[len-2] !== '/'? tag.substring(0, len-1) : tag;//normalize, if not self-terminating
		switch (tag) {

			/* 20.5.2.1  absoluteAnchor CT_AbsoluteAnchor */
			/* 20.5.2.24 oneCellAnchor */
			/* 20.5.2.33 twoCellAnchor */
			case '</xdr:absoluteAnchor':
			case '</xdr:oneCellAnchor':
			case '</xdr:twoCellAnchor':
				if(anchor){
					anchors.push(anchor);
				} else if (opts && opts.WTF) throw new Error(errorString('</'+tag.replace(/^.*?:/, '')+'>', 'anchor', match, anchor, pos));
				anchor = null;
				break;
			case '<xdr:absoluteAnchor':
			case '<xdr:oneCellAnchor':
			case '<xdr:twoCellAnchor':
				anchor = {};
				if(y.editAs){
					anchor.editAs = y.editAs;
				}
				break;


			/* CT_Blip */
			case '</a:blip':
				break;
			case '<a:blip':
				if(pic){
					pic.rId = y.embed;
					img = parse_image_ref(zip, path, pic.rId, refs, images);
					if(img){
						pic.imageId = img;
					} else if (opts && opts.WTF) throw new Error(errorString('<blip>', 'image', match, img, pos));
					img = null;
				} else if (opts && opts.WTF) throw new Error(errorString('<blip>', 'picture', match, pic, pos));
				break;

			/* 20.5.2.2 blipFill CT_BlipFillProperties */
			case '</xdr:blipFill':
			case '<xdr:blipFill':
				break;

			/* 20.5.2.7 cNvPicPr CT_NonVisualPictureProperties */
			case '</xdr:cNvPicPr':
			case '<xdr:cNvPicPr/>':
				break;
			case '<xdr:cNvPicPr':
				if(pic){
					if(y.preferRelativeResize){
						pic.preferRelativeResize = y.preferRelativeResize;
					}
				} else if (opts && opts.WTF) throw new Error(errorString('<cNvPicPr>', 'picture', match, pic, pos));
				break;


			case '</xdr:col':
				if(pos){
					pos.col = parseNum(match, 'ic');
				} else if (opts && opts.WTF) throw new Error(errorString('</col>', 'position', match, anchor, pos));
				break;
			case '<xdr:col':
				if(pos){
					pushPos(match, 'ic');
				} else if (opts && opts.WTF) throw new Error(errorString('<col>', 'position', match, anchor, pos));
				break;

			case '</xdr:colOff':
				if(pos){
					pos.colOff = parseNum(match, 'ico');
				} else if (opts && opts.WTF) throw new Error(errorString('</colOff>', 'position', match, anchor, pos));
				break;
			case '<xdr:colOff':
				if(pos){
					pushPos(match, 'ico');
				} else if (opts && opts.WTF) throw new Error(errorString('<colOff>', 'position', match, anchor, pos));
				break;

			case '</xdr:from':
				if(anchor && pos){
					anchor.from = pos;
				} else if (opts && opts.WTF) throw new Error(errorString('<from>', 'both', match, anchor, pos));
				pos = null;
				break;
			case '<xdr:from':
				pos = {};
				break;

			case '<xdr:nvPicPr':
			case '</xdr:nvPicPr':
				break;

			case '</xdr:pic':
				if(anchor){
					anchor.pic = pic;
				} else if (opts && opts.WTF) throw new Error(errorString('<pic>', 'anchor', match, anchor, pos));
				pic = null;
				break;
			case '<xdr:pic':
				pic = {};
				break;

			case '</xdr:row':
				if(pos){
					pos.row = parseNum(match, 'ir');
				} else if (opts && opts.WTF) throw new Error(errorString('</row>', 'position', match, anchor, pos));
				break;
			case '<xdr:row':
				if(pos){
					pushPos(match, 'ir');
				} else if (opts && opts.WTF) throw new Error(errorString('<row>', 'position', match, anchor, pos));
				break;

			case '</xdr:rowOff':
				if(pos){
					pos.rowOff = parseNum(match, 'iro');
				} else if (opts && opts.WTF) throw new Error(errorString('</rowOff>', 'position', match, anchor, pos));
				break;
			case '<xdr:rowOff':
				if(pos){
					pushPos(match, 'iro');
				} else if (opts && opts.WTF) throw new Error(errorString('<rowOff>', 'position', match, anchor, pos));
				break;

			case '</xdr:to':
				if(anchor && pos){
					anchor.to = pos;
				} else if (opts && opts.WTF) throw new Error(errorString('<to>', 'both', match, anchor, pos));
				pos = null;
				break;
			case '<xdr:to':
				pos = {};
				break;

			/* 20.5.2.35 wsDr CT_Drawing */
			case '<xdr:wsDr':
			case '</xdr:wsDr':
				break;

			// -------------------- unsupported/ignored -------------------------

			case '<a:ext':
			case '<a:extLst':
			case '<a:prstGeom':
			case '<a:stretch':
			case '<a:xfrm':
			case '<xdr:spPr':
				isIgnore = true;
				break;
			case '</a:ext':
			case '</a:extLst':
			case '</a:prstGeom':
			case '</a:stretch':
			case '</a:xfrm':
			case '</xdr:spPr':
				isIgnore = false;
				break;

			case '<?xml':
			case '<a14:useLocalDpi':
			case '<a:avLst/>':
			case '<a:fillRect/>':
			case '<a:off':
			case '<a:picLocks':
			case '<xdr:cNvPr':
			case '<xdr:clientData/>':
				//TODO process these yet unsupported tags/elements
				if (opts && opts.WTF) console.log('Unsupported element in wsDr element: '+tag);
				break;
			default:
				if (opts && opts.WTF) throw new Error('Unrecognized ' + y[0] + ' in wsDr');
		}
	}

	tagregex.lastIndex = 0;
	return {"!type": "drawing", anchors: anchors};
}

/* FEATURE[russa]: parse image ref */
function parse_image_ref(zip, path, rId, imgRefs, images){
	for(var n in imgRefs){
		if(imgRefs[n].Id === rId){
			var uid = n;
			if(images[uid]){
				return uid;
			}
			var target = imgRefs[n].Target;
			var imgfile = resolve_path(target, path);
			if(imgfile){
				var data = getdatabin(getzipfile(zip, imgfile));
				images[uid] = {id: uid, data: data};
				return uid;
			}
		}
	}
}
