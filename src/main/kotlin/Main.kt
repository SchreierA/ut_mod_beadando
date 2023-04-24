import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.chocosolver.solver.Model
import org.chocosolver.solver.variables.IntVar
import java.io.File

val ROUTE_NAMES = Pair("Vá.1->Vá.2", "Vá.2->Vá.1")
val ROUTE_NAMES_CLEARED = Pair("Vá.1toVá.2", "Vá.2toVá.1")
const val DRAW_LIMIT = 10

fun stringifyEdgeList(edges: List<String>, keepCommas: Boolean = false): String {
    return edges.toString()
        .replace("[", "")
        .replace("]", "")
        .replace(",", if(keepCommas) ",\n" else "\n")
}

fun loadWorkbook(): Workbook {
    File("src/input/Input_VSP.xlsx").let {
        return XSSFWorkbook(it)
    }
}

fun processWorkbook(workbook: Workbook): Map<String, List<Pair<Double, Double>>> {
    val r1t2 = mutableListOf<Pair<Double, Double>>()
    val r2t1 = mutableListOf<Pair<Double, Double>>()
    val mapRowToEmbarkArrivalPair = { row: Row, offset: Double ->
        Pair(row.getCell(0).numericCellValue + offset, row.getCell(1).numericCellValue + offset)
    }
    val offsets1 = mutableListOf<Pair<Double, Double>>()
    val offsets2 = mutableListOf<Pair<Double, Double>>()
    workbook.getSheetAt(0).let {
        arrayOf(Pair(8, 33), Pair(34, 63)).forEach { bounds ->
            for (i in bounds.first..bounds.second) {
                it.getRow(i)?.let { offsets ->
                    val till = offsets.getCell(7)?.localDateTimeCellValue ?: return@let
                    offsets1.add(
                        Pair(
                            till.hour * 60.0 + till.minute,
                            offsets.getCell(8).numericCellValue + offsets.getCell(9).numericCellValue
                        )
                    )
                }
            }
        }
    }
    offsets1.reverse()
    offsets2.reverse()
    workbook.getSheetAt(1)?.let { sheet ->
        sheet.forEachIndexed { index, row ->
            if(index == 0) return@forEachIndexed
            if(row.getCell(2).toString() == ROUTE_NAMES.first) {
                r1t2.add(mapRowToEmbarkArrivalPair(row,
                    offsets1.find { offset -> offset.first < row.getCell(0).numericCellValue }?.second ?: 0.0))
                return@forEachIndexed
            }
            r2t1.add(mapRowToEmbarkArrivalPair(row,
                offsets2.find{offset -> offset.first < row.getCell(0).numericCellValue}?.second ?: 0.0))
        }
    }
    return mapOf(ROUTE_NAMES.first to r1t2, ROUTE_NAMES.second to r2t1)
}

fun buildCompatibilityGraph(routes: Map<String, List<Pair<Double, Double>>>): List<String> {
    val edges = mutableListOf<String>()
    routes[ROUTE_NAMES.first]?.forEachIndexed { index, current ->
        if(index > DRAW_LIMIT) return@forEachIndexed
        routes[ROUTE_NAMES.second]?.forEachIndexed { innerIndex, destination ->
            if(innerIndex > DRAW_LIMIT) return@forEachIndexed
            if(current.second < destination.first)
                edges.add("\n${ROUTE_NAMES_CLEARED.first}::$index --- ${ROUTE_NAMES_CLEARED.second}::$innerIndex")
        }
    }
    routes[ROUTE_NAMES.second]?.forEachIndexed { index, current ->
        if(index > DRAW_LIMIT) return@forEachIndexed
        routes[ROUTE_NAMES.first]?.forEachIndexed { innerIndex, destination ->
            if(innerIndex > DRAW_LIMIT) return@forEachIndexed
            if(current.second < destination.second)
                edges.add("\n${ROUTE_NAMES_CLEARED.second}::$index -.- ${ROUTE_NAMES_CLEARED.first}::$innerIndex")
        }
    }
    return edges
}

fun buildConnectionBaseMultiCommodityGraph(routes: Map<String, List<Pair<Double, Double>>>): List<String> {
    val edges = mutableListOf<String>()
    routes[ROUTE_NAMES.first]?.forEachIndexed { index, current ->
        if(index > DRAW_LIMIT) return@forEachIndexed
        edges.add("\n${ROUTE_NAMES_CLEARED.first}::$index -.- dep")
        routes[ROUTE_NAMES.second]?.forEachIndexed { innerIndex, destination ->
            if(innerIndex > DRAW_LIMIT) return@forEachIndexed
            if(current.second < destination.first)
                edges.add("\n${ROUTE_NAMES_CLEARED.first}::$index --> ${ROUTE_NAMES_CLEARED.second}::$innerIndex")
        }
    }
    routes[ROUTE_NAMES.second]?.forEachIndexed { index, current ->
        if(index > DRAW_LIMIT) return@forEachIndexed
        edges.add("\n${ROUTE_NAMES_CLEARED.second}::$index -.- dep")
        routes[ROUTE_NAMES.first]?.forEachIndexed { innerIndex, destination ->
            if(innerIndex > DRAW_LIMIT) return@forEachIndexed
            if(current.second < destination.second)
                edges.add("\n${ROUTE_NAMES_CLEARED.second}::$index --> ${ROUTE_NAMES_CLEARED.first}::$innerIndex")
        }
    }
    return edges
}

fun writeGraph(edges: List<String>, fileName: String, templateUrl: String, keepCommas: Boolean = false) {
    val graph = stringifyEdgeList(edges, keepCommas)
    var template = ""
    File(templateUrl).let {
        template = it.readText().replace("#####replace#####", graph)
    }
    if (template.isEmpty()) return
    val file = File("output/$fileName.html")
    if (file.exists()) file.writeText("")
    /*val writer = BufferedWriter(FileWriter(file))
    writer.apply {
        write(template)
        flush()
    }*/
    file.writeText(template)
    println("$fileName file created")
}

fun buildTimeSpaceGraphNodes(routes: Map<String, List<Pair<Double, Double>>>): String {
    var graph = ""
    routes.forEach { station ->
        graph += "<div class=\"place\"><p class=\"title\">${station.key.substringAfter("->")}</p>\n"
        station.value.forEachIndexed { index, route ->
            graph += "<div id=\"${station.key}::$index\" class=\"node\">${route.first}</div>"
        }
        graph += "</div>\n"
    }
    graph +="<div class=\"place\"><p class=\"title\">Depo</p><div id=\"depo\" class=\"node\">Depo</div></div>"
    return graph
}

fun buildTimeSpaceEdges(routes: Map<String, List<Pair<Double, Double>>>): String {
    var script = "<script>"
    routes.forEach { station ->
        station.value.forEachIndexed { index, route ->
            if(index > DRAW_LIMIT) return@forEachIndexed
            script += "" +
                    "new LeaderLine(\n" +
                    "  document.getElementById('depo'),\n" +
                    "  document.getElementById('${station.key}::$index'),\n" +
                    " {dash: true});"
            script += "" +
                    "new LeaderLine(\n" +
                    "  document.getElementById('${station.key}::$index'),\n" +
                    "  document.getElementById('depo'),\n" +
                    " {dash: true});"
            routes.forEach { s ->
                if (s.key == station.key) return@forEach
                s.value.forEachIndexed { j, r ->
                    if(j > DRAW_LIMIT) return@forEachIndexed
                    if (route.second < r.first)
                    script += "" +
                            "new LeaderLine(\n" +
                            "  document.getElementById('${station.key}::$index'),\n" +
                            "  document.getElementById('${s.key}::$j')\n" +
                            ");"
                }
            }
        }
    }
    script += "</script>"
    return script
}

fun buildTimeSpaceGraph(routes: Map<String, List<Pair<Double, Double>>>): String {
    var graph = buildTimeSpaceGraphNodes(routes)
    graph += buildTimeSpaceEdges(routes)
    return graph
}

fun buildTimeSpaceModel(routes: Map<String, List<Pair<Double, Double>>>): Model {
    val model = Model("Bus schedule")

    val routesInModel = mutableListOf<IntVar>()
    model.intVar("depot", 0, 999999)
    routes.forEach { station ->
        station.value.forEachIndexed { index, route ->
            routesInModel.add(
                model.intVar("${station.key}::$index", route.first.toInt(), route.second.toInt())
            )
        }
    }
    /*for (i in 0 until routesInModel.size) {
        for (j in i+1 until routesInModel.size) {
            val c1 = model.arithm(routesInModel[i], "<", routesInModel[j])
            val c2 = model.arithm(routesInModel[j], "<", routesInModel[i])
            c1.implies(c2.reify())
        }
    }*/

    val numBuses = model.intVar(1, routesInModel.size)
    val busesUsed = model.intVar(0, routesInModel.size)
    for (route in routesInModel) {
        model.ifThen(model.arithm(route, ">", 0), model.arithm(busesUsed, "+", model.intVar(1), ">=", numBuses))
    }
    model.setObjective(Model.MINIMIZE, numBuses)

    return model
}

fun lpifyChocoModelHead(model: Model): String {
    // Define the LP model
    var lpModel = "Maximize\n"
    lpModel += "obj: "
    lpModel += model.objective
    lpModel += "\n\nSubject To\n"
    model.cstrs.forEach {
        lpModel += "${it}\n"
    }
    lpModel += "\nBounds\n"
    model.vars.forEach {
        lpModel += "${it.name}: ${it.asIntVar().lb}\n"
        lpModel += "${it.name}: ${it.asIntVar().ub}\n"
    }
    lpModel += "\nEnd\n\n"
    return lpModel
}

fun lpifyChocoModelTail(model: Model): String {
    // Map Choco Solver variables to LP variables
    val lpVars = DoubleArray(model.vars.size)
    for (i in model.vars.indices) {
        lpVars[i] = model.vars[i].asIntVar().value.toDouble()
    }

    var res = ""
    res +=  "Solution\n"
    for (i in lpVars.indices) {
        res += "x[$i]=${lpVars[i]}\n"
    }
    return res
}


fun solveModel(model: Model) {
    var output = lpifyChocoModelHead(model)
    val solver = model.solver
    if (solver.solve()) {
        output += lpifyChocoModelTail(model)
        println("Solution found: ${ solver.findSolution()}")
    } else {
        println("No solution found.")
    }
    val file = File("output/model.lp")
    file.writeText(output)
}

fun main(args: Array<String>) {
    val workbook = loadWorkbook()
    val routes = processWorkbook(workbook)

    //val compatibilityGraph = buildCompatibilityGraph(routes)
    //writeCompatibilityGraph(compatibilityGraph)

    val connectionBaseMultiCommodityGraph = buildConnectionBaseMultiCommodityGraph(routes)
    writeGraph(connectionBaseMultiCommodityGraph, "connection", "templates/template.html")

    val timeSpaceGraph = buildTimeSpaceGraph(routes)
    writeGraph(listOf(timeSpaceGraph), "ts", "templates/empty-template.html", true)

    val tsModel = buildTimeSpaceModel(routes)
    solveModel(tsModel)
}