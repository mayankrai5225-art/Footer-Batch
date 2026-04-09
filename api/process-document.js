module.exports = async (req, res) => {
	const mod = await import("../frontend/api/process-document.js");
	return mod.default(req, res);
};