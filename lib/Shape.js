import { PowerPointObject} from './proxy.js';
import { Fill, Line} from './style.js';
import { TextFrame2} from './TextBox.js';

export class Shape extends PowerPointObject {

	#ticket;

	constructor( ticket) {
		super( ticket);
		this.#ticket = ticket;
	}

	set fillColor( rgb) {
		const bgr = ( rgb & 0x0000FF) << 16 | ( rgb & 0x00FF00) | ( rgb & 0xFF0000) >>> 16;
		this.#Fill.Solid();
		this.#Fill.ForeColor.RGB = bgr;
	}

	get lineStyle() {
		return this.#Line.Visible === 0? undefined: this.#Line.DashStyle;
	}

	set lineStyle( style) {
		if( style === undefined)
			this.#Line.Visible = 0;
		else {
			this.#Line.Visible = -1;
			this.#Line.DashStyle = style;
		}
	}

	get lineWeight() {
		return this.#Line.Weight;
	}

	set lineWeight( weight) {
		this.#Line.Weight = weight;
	}

	get #Fill() {
		return this.#ticket.get( 'Fill', ticket => new Fill( ticket));
	}

	get #Line() {
		return this.#ticket.get( 'Line', ticket => new Line( ticket));
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
