export class Background {

	#ticket;

	constructor( ticket) {
		this.#ticket = ticket;
	}

	get Fill() {
		return this.#ticket.get( 'Fill', ticket => new Fill( ticket));
	}
}

export class Fill {

	#ticket;

	constructor( ticket) {
		this.#ticket = ticket;
	}

	get ForeColor() {
		return this.#ticket.get( 'ForeColor', ticket => new ForeColor( ticket));
	}

	Solid() {
		this.#ticket.get( 'Solid', ticket => () => ticket.apply( this, []))();
	}

	set Visible( visible) {
		this.#ticket.set( 'Visible', visible);
	}
}

export class ForeColor {

	#ticket;

	constructor( ticket) {
		this.#ticket = ticket;
	}

	set RGB( bgr) {
		this.#ticket.set( 'RGB', bgr);
	}
}

export class Line {

	#ticket;

	constructor( ticket) {
		this.#ticket = ticket;
	}

	get DashStyle() {
		return this.#ticket.get( 'DashStyle');
	}

	set DashStyle( style) {
		this.#ticket.set( 'DashStyle', style);
	}

	get Weight() {
		return this.#ticket.get( 'Weight');
	}

	set Weight( weight) {
		this.#ticket.set( 'Weight', weight);
	}

	get Visible() {
		return this.#ticket.get( 'Visible');
	}

	set Visible( visible) {
		this.#ticket.set( 'Visible', visible);
	}

	get EndArrowheadStyle() {
		return this.#ticket.get( 'EndArrowheadStyle');
	}

	set EndArrowheadStyle( style) {
		this.#ticket.set( 'EndArrowheadStyle', style);
	}
}
