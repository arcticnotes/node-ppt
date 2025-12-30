import { PowerPointObject} from './proxy.js';
import { Fill } from './style.js';

export class TextBox extends PowerPointObject {

	#ticket;

	constructor( ticket) {
		super( ticket);
		this.#ticket = ticket;
	}

	get name() {
		return this.#ticket.get( 'Name');
	}

	set name( name) {
		this.#ticket.set( 'Name', name);
	}

	get content() {
		return this.#TextFrame2.TextRange.Text;
	}

	set content( content) {
		this.#TextFrame2.TextRange.Text = content;
	}

	get tracking() {
		return this.#TextFrame2.TextRange.Font.Spacing;
	}

	set tracking( trakcing) {
		this.#TextFrame2.TextRange.Font.Spacing = trakcing;
	}

	get fontSize() {
		return this.#TextFrame2.TextRange.Font.Size;
	}

	set fontSize( size) {
		this.#TextFrame2.TextRange.Font.Size = size;
	}

	get textColor() {
		const bgr = this.#TextFrame2.TextRange.Font.Fill.ForeColor.RGB;
		return ( bgr & 0x0000FF) << 16 | ( bgr & 0x00FF00) | ( bgr & 0xFF0000) >>> 16;
	}

	set textColor( rgb) {
		const bgr = ( rgb & 0x0000FF) << 16 | ( rgb & 0x00FF00) | ( rgb & 0xFF0000) >>> 16;
		this.#TextFrame2.TextRange.Font.Fill.ForeColor.RGB = bgr;
	}

	get #TextFrame2() {
		return this.#ticket.get( 'TextFrame2', ticket => new TextFrame2( ticket));
	}
}

export class TextFrame2 {

	#ticket;

	constructor( ticket) {
		this.#ticket = ticket;
	}

	get TextRange() {
		return this.#ticket.get( 'TextRange', ticket => new TextRange( ticket));
	}
}

class TextRange {

	#ticket;

	constructor( ticket) {
		this.#ticket = ticket;
	}

	get Text() {
		return this.#ticket.get( 'Text');
	}

	set Text( text) {
		this.#ticket.set( 'Text', text);
	}

	get Font() {
		return this.#ticket.get( 'Font', ticket => new Font( ticket));
	}
}

class Font {

	#ticket;

	constructor( ticket) {
		this.#ticket = ticket;
	}

	get Fill() {
		return this.#ticket.get( 'Fill', ticket => new Fill( ticket));
	}

	get Spacing() {
		return this.#ticket.get( 'Spacing');
	}

	set Spacing( spacing) {
		this.#ticket.set( 'Spacing', spacing);
	}

	get Size() {
		return this.#ticket.get( 'Size');
	}

	set Size( size) {
		this.#ticket.set( 'Size', size);
	}
}
