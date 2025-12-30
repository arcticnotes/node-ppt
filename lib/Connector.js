import { PowerPointObject} from './proxy.js';
import { Fill, Line} from './style.js';

export class Connector extends PowerPointObject {

	#ticket;

	constructor( ticket) {
		super( ticket);
		this.#ticket = ticket;
	}

	get endStyle() {
		return this.#Line.EndArrowheadStyle;
	}

	set endStyle( style) {
		this.#Line.EndArrowheadStyle = style;
	}

	get #Line() {
		return this.#ticket.get( 'Line', ticket => new Line( ticket));
	}

	connectFrom( shape, port) {
		this.#ConnectorFormat.BeginConnect( shape, port);
	}

	connectTo( shape, port) {
		this.#ConnectorFormat.EndConnect( shape, port);
	}

	get #ConnectorFormat() {
		return this.#ticket.get( 'ConnectorFormat', ticket => new ConnectFormat( ticket));
	}
}

class ConnectFormat extends PowerPointObject {

	#ticket;

	constructor( ticket) {
		super( ticket);
		this.#ticket = ticket;
	}

	BeginConnect( shape, port) {
		this.#ticket.get( 'BeginConnect', ticket => ( shape, port) => ticket.apply( this, [ shape, port]))( shape, port);
	}

	EndConnect( shape, port) {
		this.#ticket.get( 'EndConnect', ticket => ( shape, port) => ticket.apply( this, [ shape, port]))( shape, port);
	}
}
