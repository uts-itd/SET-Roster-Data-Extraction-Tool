class Staff {
	#deployments;	// array of deployment objects

	constructor(name) {
		this.name = name;
	}

	addDeployment(deployment) {
		this.#deployments.push(deployment);
	}
}

module.exports = Staff;
