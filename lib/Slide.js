import { Connector} from './Connector.js';
import { PowerPointObject} from './proxy.js';
import { Shape} from './Shape.js';
import { Background} from './style.js';
import { TextBox} from './TextBox.js';

export class Slide extends PowerPointObject {

	#ticket;

	constructor( ticket) {
		super( ticket);
		this.#ticket = ticket;
	}

	get index() {
		return this.#ticket.get( 'SlideIndex');
	}

	get name() {
		return this.#ticket.get( 'Name');
	}

	set name( name) {
		this.#ticket.set( 'Name', name);
	}

	set background( rgb) {
		const bgr = ( rgb & 0x0000FF) << 16 | ( rgb & 0x00FF00) | ( rgb & 0xFF0000) >>> 16;
		const background = this.#ticket.get( 'Background', ticket => new Background( ticket));
		this.#ticket.set( 'FollowMasterBackground', false);
		background.Fill.ForeColor.RGB = bgr;
		background.Fill.Solid();
		background.Fill.Visible = -1;
	}

	newTextBox( orientation, left, top, width, height) {
		return this.#Shapes.AddTextBox( orientation, left, top, width, height);
	}

	newShape( type, left, top, width, height) {
		return this.#Shapes.AddShape( type, left, top, width, height);
	}

	newConnector( type, left, top, width, height) {
		return this.#Shapes.AddConnector( type, left, top, width, height);
	}

	get #Shapes() {
		return this.#ticket.get( 'Shapes', ticket => new Shapes( ticket));
	}
}

class Shapes {

	#ticket;

	constructor( ticket) {
		this.#ticket = ticket;
	}

	AddTextBox( orientation, left, top, width, height) {
		return this.#ticket.get( 'AddTextBox', ticket => ( ...args) => ticket.apply( this, args, t => new TextBox( t)))(
				orientation, left, top, width, height);
	}

	AddShape( type, left, top, width, height) {
		return this.#ticket.get( 'AddShape', ticket => ( ...args) => ticket.apply( this, args, t => new Shape( t)))(
				type, left, top, width, height);
	}

	AddConnector( type, left, top, width, height) {
		return this.#ticket.get( 'AddConnector', ticket => ( ...args) => ticket.apply( this, args, t => new Connector( t)))(
				type, left, top, width, height);
	}
}
