import sys
from functools import lru_cache
import powerfactory as pf

sys.path.append(r'\\Ecasd01\WksMgmt\PowerFactory\Scripts\pfTextOutputs')
import pftextoutputs

sys.path.append(r'\\Ecasd01\WksMgmt\PowerFactory\Scripts\pfAttributeManipulator')

import logging

logger = logging.getLogger('Subtrans Graphical Connection')


def run_main():
    """Get app and setup logging"""

    app = pf.GetApplication()
    main(app)


def main(app):
    with pftextoutputs.PowerFactoryLogging(
        pf_app=app,
        add_handler=True,
        handler_level=logging.DEBUG,
        formatter=pftextoutputs.PFFormatter(
            '%(module)s: %(funcName)s: Line: %(lineno)d: %(message)s'
        ),
    ) as pflogger:
        app.PrintInfo(f'Manual Print using app not {logger}')
        logger.info(f'Manual Print using {logger}')

        active_project = app.GetActiveProject()
        if not active_project:
            logging.error(
                'No Active Project, ' 'Script Stopped before it does anything'
            )
            return

        current_script = app.GetCurrentScript()

        try:
            revert = current_script.add_reversion
        except AttributeError:
            revert = False
        if revert:
            add_reversion(app)

        try:
            max_elms = current_script.max_elms
        except AttributeError:
            max_elms = 10 ** 100
            logger.warning(
                f'{current_script} does not have a "max_elms"'
                f'attribute, assuming {max_elms}'
            )

        find_undrawn_elements(app, max_elms)

        return


def find_undrawn_elements(app, max_elms=10 ** 100):
    graphics = app.GetActiveProject().GetContents('*.IntGrfnet', True)

    for graphic in graphics:
        draw_undrawn_branches_in_diagram(graphic=graphic, max_elms=max_elms)


def draw_undrawn_branches_in_diagram(graphic, max_elms=10 ** 100):
    if graphic.GetAttribute('iFrzPerm'):
        # Skip a non perminant diagram or one that is write protected
        logger.debug(f'Skipping {graphic} as it is write protected')
        return
    logger.debug(graphic)

    objs_in_graphic = [
        gr.GetAttribute('pDataObj') for gr in graphic.GetContents('*.IntGrf')
    ]
    # Add in children of container objects to handle branches.
    grouping_object_cn = ['ElmBranch', 'ElmSubstat', 'ElmTrfstat', 'ElmSite']
    for o in objs_in_graphic:
        if o and o.GetClassName() in grouping_object_cn:
            objs_in_graphic.extend(o.GetContents('*.Elm*'))

    if len(objs_in_graphic) > max_elms:
        logger.warning(f'Skipping {graphic} as it has too many elements')
    terms_in_graphic = list()

    non_terms_in_graphic = list()
    for o in objs_in_graphic:
        if o:
            if o.GetClassName() == 'ElmTerm':
                terms_in_graphic.append(o)
            else:
                non_terms_in_graphic.append(o)

    logger.info(
        f"Checking for graphic representations to add to "
        f"{graphic}'s {len(terms_in_graphic)}"
    )

    for term in terms_in_graphic:
        term_cubs = term.GetAttribute('cubics')
        objs_assos_term = (c.GetAttribute('obj_id') for c in term_cubs)
        objs_assos_term = [o for o in objs_assos_term if o]
        for obj in objs_assos_term:
            if obj not in get_drawn_objects_in_diagram(graphic):
                # Draw the undrawn object
                connection_num = determine_connection_number(
                    obj=obj, associated_term=term
                )
                add_to_diagram(
                    new_object=obj,
                    terminal=term,
                    diagram=graphic,
                    connection_num=connection_num,
                )

    return


def get_drawn_objects_in_diagram(diagram):
    terms_in_graphic = (
        gr.GetAttribute('pDataObj') for gr in diagram.GetContents('*.IntGrf')
    )
    return [o for o in terms_in_graphic if o]


def add_reversion(app):
    logger.warning('Attempting to revert the project to latest version')
    active_project = app.GetActiveProject()

    if active_project:
        latest_version = active_project.GetLatestVersion(0)

        if latest_version:
            logger.warning(f'Reverting {active_project} to {latest_version}')
            logger.info('Deactivating')
            active_project.Deactivate()
            logger.info(f'Rolling Back to {latest_version}')
            latest_version.Rollback()
            logger.info(f'Activating {active_project}')
            active_project.Activate()
            logger.info(f'Activated project')


def determine_connection_number(obj, associated_term):
    """
    Determine which connection number the cubical between the
    terminal and object is.
    Raises a ValueError if they don't match
    """

    for cub in associated_term.GetAttribute('cubics'):
        cub_obj = cub.GetAttribute('obj_id')
        if cub_obj == obj:
            try:
                return cub.GetAttribute('obj_bus')
            except AttributeError:
                logger.error(f'Unable to get {cub}.obj_bus')
    logger.error(f'{obj} is not directly connected with {associated_term}')
    raise ValueError(f'{obj} is not directly connected with {associated_term}')


def add_to_diagram(new_object, terminal, diagram, connection_num=0):
    """

    :param new_object: The object to be drawn
    :param terminal: The Terminal to draw it connected to
    :param diagram: The Diagram to draw it in
    :param connection_num:
    :param new_obj_dia:
    :return:
    """
    logger.info(
        f'Adding a diagram object for {new_object} '
        f'associated with {terminal} in {diagram}'
    )

    term_dia = None

    for dia_obj in diagram.GetContents('*.IntGrf'):
        ass_obj = dia_obj.GetAttribute('pDataObj')
        # logger.debug(f'{dia_obj} is associated with {ass_obj}')
        if ass_obj and ass_obj == terminal:
            term_dia = dia_obj

    if not term_dia:
        logger.warning(
            f'Unable to create diagram object for {new_object} '
            f'as {terminal} diagram object was not found'
        )
        return

    up_or_down = determine_y_diff_orientation(
        terminal=terminal, term_dia=term_dia, diagram=diagram
    )

    if up_or_down is None:
        # if the orientation is not clear if there is another terminal
        # in the diagram it can be drawn from draw it from that

        other_terms = list()

        for connection in range(5):
            con_cub = new_object.GetCubicle(connection)
            if con_cub:
                con_term = con_cub.GetAttribute('cterm')
                if con_term and con_term != terminal:
                    other_terms.append(con_term)

        diagram_objs = get_drawn_objects_in_diagram(diagram=diagram)
        if any(other_term in diagram_objs for other_term in other_terms):
            logger.warning(
                f'Not Creating a diagram for {new_object} based'
                f' on its connection to {terminal} as that terminal'
                f' is horizontally drawn and does not indicate'
                f' which side the new picture should be drawn on'
            )

            return
        else:
            up_or_down = -1 if up_or_down else 1

    new_obj_dia = diagram.CreateObject('IntGrf', new_object.GetAttribute('loc_name'))
    logger.debug(f'Created {new_obj_dia} \t ' f'{new_obj_dia.GetFullName()}')

    new_obj_dia.SetAttribute('pDataObj', new_object)

    new_obj_dia.SetAttribute('sSymNam', get_default_symbol_type(new_object))

    term_x = term_dia.GetAttribute('rCenterX')
    term_y = term_dia.GetAttribute('rCenterY')
    logger.debug(f'{terminal} / {term_dia} is centred on x: {term_x}, y: {term_y}')

    # Find the best spot to put the new line by putting it in the largest gap
    # on the terminal
    ideal_x_location = determine_terminal_x_location(
        terminal, term_dia, diagram, up_or_down=up_or_down
    )

    y_difference = 10 * up_or_down * -1

    new_obj_dia.SetAttribute('rCenterX', ideal_x_location)
    new_obj_dia.SetAttribute('rCenterY', term_y + y_difference)

    new_obj_connection = new_obj_dia.CreateObject('IntGrfcon', 'GCO')

    logger.debug(
        f'Created {new_obj_connection} \t ' f'{new_obj_connection.GetFullName()}'
    )

    obj_con_x, obj_con_y = get_connection_x_y(
        new_obj_dia=new_obj_dia, connection_num=connection_num
    )

    rx = new_obj_connection.GetAttribute('rX')
    ry = new_obj_connection.GetAttribute('rY')

    rx[0] = ideal_x_location
    rx[1] = ideal_x_location

    ry[0] = term_y + y_difference
    ry[1] = term_y

    new_obj_connection.SetAttribute('rX', rx)
    new_obj_connection.SetAttribute('rY', ry)
    new_obj_connection.SetAttribute('iDatConNr', connection_num)
    new_obj_connection.SetAttribute('rLinWd', 0.1)

    logger.debug(f'New Graphic Object {new_obj_dia} created in {diagram}')

    # Check if the other end needs to be connected
    create_other_diagram_connections(
        obj=new_object,
        new_obj_dia=new_obj_dia,
        connected_term=terminal,
        diagram=diagram,
    )


def create_other_diagram_connections(obj, new_obj_dia, connected_term, diagram):
    """
    Check if any other terminals are already drawn in the diagram,
    if so connect them to the object
    """

    obj_x_centre = new_obj_dia.GetAttribute('rCenterX')
    obj_y_centre = new_obj_dia.GetAttribute('rCenterY')

    for connection_number in range(5):
        cub = obj.GetCubicle(connection_number)
        if cub:
            term = cub.GetAttribute('cterm')
            if term:
                if term == connected_term:
                    continue

                term_dia = get_obj_graphic_from_diagram(term, diagram)
                if term_dia is None:
                    # If the terminal isn't drawn in this diagram,
                    # adding the connection is useless
                    continue
                term_y = term_dia.GetAttribute('rCenterY')

                # Find the best spot to put the new line by putting it in the largest gap
                # on the terminal
                up_or_down = determine_y_diff_orientation(
                    terminal=term, term_dia=term_dia, diagram=diagram
                )
                # No check for None ness here as the original
                # orientation has been decided

                ideal_x_location = determine_terminal_x_location(
                    term, term_dia, diagram, up_or_down=up_or_down
                )

                obj_con_x, obj_con_y = get_connection_x_y(
                    new_obj_dia=new_obj_dia, connection_num=connection_number
                )

                create_connection(
                    new_obj_dia=new_obj_dia,
                    x1=obj_x_centre,
                    y1=obj_y_centre,
                    x2=ideal_x_location,
                    y2=term_y,
                    connection_num=connection_number,
                )


def get_connection_x_y(new_obj_dia, connection_num):
    """Use the connection point numbers associated with the symbol"""

    obj_x_centre = new_obj_dia.GetAttribute('rCenterX')
    obj_y_centre = new_obj_dia.GetAttribute('rCenterY')
    obj_x_scale = new_obj_dia.GetAttribute('rSizeX')
    obj_y_scale = new_obj_dia.GetAttribute('rSizeY')
    obj_sym_name = new_obj_dia.GetAttribute('sSymNam')

    obj_sym_full_name = f'\\Sys\\Lib\\Grf\\Symbols\\SGL\\{obj_sym_name}.IntSym'
    obj_sym = new_obj_dia.SearchObject(obj_sym_full_name)

    if not obj_sym:
        raise RuntimeError(f'Unable to find {obj_sym_full_name}')

    sym_xs = obj_sym.GetAttribute('rX')
    sym_ys = obj_sym.GetAttribute('rY')

    connection_x_delta = sym_xs[connection_num]
    connection_y_delta = sym_ys[connection_num]

    new_x = obj_x_centre + connection_x_delta * obj_x_scale
    new_y = obj_y_centre + connection_y_delta * obj_y_scale

    return new_x, new_y


def create_connection(new_obj_dia, x1, y1, x2, y2, connection_num):
    new_obj_connection = new_obj_dia.CreateObject('IntGrfcon', 'GCO')

    logger.debug(
        f'Created {new_obj_connection} \t ' f'{new_obj_connection.GetFullName()}'
    )

    rx = new_obj_connection.GetAttribute('rX')
    ry = new_obj_connection.GetAttribute('rY')

    rx[0] = x1
    rx[1] = x2

    ry[0] = y1
    ry[1] = y2

    new_obj_connection.SetAttribute('rX', rx)
    new_obj_connection.SetAttribute('rY', ry)
    new_obj_connection.SetAttribute('iDatConNr', connection_num)
    new_obj_connection.SetAttribute('rLinWd', 0.1)

    return new_obj_connection


symbol_defaults = {
    'ElmTerm': 'TermStrip',
    'ElmLne': 'd_lin',
    'ElmCoup': 'd_couple',
    'ElmShnt': 'd_shunt',
    'ElmVac': 'd_vac',
    'StaSua': 'd_sua',
    'ElmTr2': 'd_tr2',
    'ElmTr3': 'd_tr3',
    'ElmLod': 'd_load',
    'ElmNec': 'd_nec',
    'ElmXnet': 'd_net',
    'ElmSym': 'd_symg',
}


def get_default_symbol_type(obj):
    class_name = obj.GetClassName()

    try:
        symbol_name = symbol_defaults[class_name]
    except KeyError:
        logger.error(f'No Default Symbol Provided for {class_name}')
        return 'TermStrip'
    return symbol_name


# Cache whether it is up or down to avoid up then down then up then down
@lru_cache(2 * 8)
def determine_y_diff_orientation(terminal, term_dia, diagram):
    """
    Attempt to determine if most of the existing graphics are up or down

    If up return 1 to indicate it should be drawn down,
    else return -1 to indicate it should be drawn up
    """

    obj_ys = list()
    element_orientation = list()
    term_y = term_dia.GetAttribute('rCenterY')

    term_objs = list()
    for cubical in terminal.GetAttribute('cubics'):
        cub_xs, cub_ys = get_xs_ys(cubical, diagram)
        obj_ys.append(cub_ys)
        logger.debug(f'Found {cubical} with ys {cub_ys}')
        non_zero_ys = [y for y in cub_ys if y]
        if not non_zero_ys:
            element_orientation.append(0)
            continue

        min_cub_y = min(non_zero_ys)
        max_cub_y = max(non_zero_ys)
        logger.debug(f'{cubical} has the min ys as {min_cub_y}, max ys as {max_cub_y}')

        if min_cub_y == max_cub_y and min_cub_y == term_y:
            return None
        elif min_cub_y >= term_y:
            element_orientation.append(1)
        elif max_cub_y <= term_y:
            element_orientation.append(-1)
        elif max_cub_y - term_y > term_y - min_cub_y:
            element_orientation.append(1)
        else:
            element_orientation.append(-1)

    logger.debug(f'{terminal} elm orientation: {element_orientation}')

    up_or_down = 1 if sum(element_orientation) >= 0 else -1

    logger.debug(f'Determined that {terminal} is mostly {up_or_down}')

    return up_or_down


def get_obj_graphic_from_diagram(obj, diagram):
    """
    Return the IntGrf object for an Elm Obj in Diagram.
    returns None if not found
    """

    grf_objs = diagram.GetContents('*.IntGrf')
    for gr in grf_objs:
        o = gr.GetAttribute('pDataObj')
        if o == obj:
            return gr

    return None


def determine_terminal_x_location(
    terminal, term_dia, diagram, up_or_down, include_results=False
):
    """
    Ensure that the new object is put in the largest gap in the terminal sets

    if up = 1 find the larger top gab, otherwise largest under gap

    """
    up_side = up_or_down != 1

    terminal_length = 4.375
    normal_size_scale = 3
    results_size = 6 if include_results else 0

    symbol_name = term_dia.GetAttribute('sSymNam')
    term_x = term_dia.GetAttribute('rCenterX')
    term_y = term_dia.GetAttribute('rCenterY')
    term_x_scale = term_dia.GetAttribute('rSizeX')

    if symbol_name in ['TermStrip', 'TermStripThin']:
        start = (
            term_x - (terminal_length * normal_size_scale * term_x_scale) + results_size
        )
        end = term_x + (terminal_length * normal_size_scale * term_x_scale)

    elif symbol_name in ['ShortTermStrip']:
        start = term_x - (terminal_length * term_x_scale) - results_size
        end = term_x + (terminal_length * term_x_scale)

    elif symbol_name in ['PointTerm']:
        start = term_x
        end = term_x

    else:
        logger.error(f'{symbol_name} not listed')
        raise RuntimeError(f'{symbol_name} not listed')

    logger.debug(f'start: {start} end: {end}')
    xs_that_matter = list()

    xs_that_matter.append(start)
    for cubical in terminal.GetAttribute('cubics'):
        cub_xs = get_relevant_xs(cubical, diagram, y_t=term_y, up_side=up_side)

        new_xs = [x for x in cub_xs if start <= x and x <= end]

        logger.debug(f'adding {new_xs} associated with {cubical}, up_side = {up_side}')
        xs_that_matter.extend(new_xs)

    xs_that_matter.append(end)

    if not xs_that_matter:
        raise RuntimeError(f'Somehow the exs that matter are empty ')

    x_spot = get_centre_of_largest_gap(xs_that_matter)
    logger.debug(f'Selected {x_spot} from {xs_that_matter}')
    return x_spot


def get_xs_ys(cubical, diagram):
    obj = cubical.GetAttribute('obj_id')

    xs = list()
    ys = list()

    connection_id = cubical.GetAttribute('obj_bus')
    for dia_obj in diagram.GetContents('*.IntGrf'):
        other_object = dia_obj.GetAttribute('pDataObj')
        if other_object and other_object == obj:
            break
    else:
        # no diagram found return None
        logger.debug(f'{obj}, {cubical}')
        return xs, ys

    x_center = dia_obj.GetAttribute('rCenterX')
    y_center = dia_obj.GetAttribute('rCenterY')
    xs.append(x_center)
    ys.append(y_center)

    logger.debug(
        f'{obj} associated with {cubical} and {dia_obj} '
        f'has centers on {x_center}, {y_center}'
    )

    for connection_object in dia_obj.GetContents('*.IntGrfcon'):
        connection_xs = connection_object.GetAttribute('rX')
        valid_xs = [x for x in connection_xs if x != -1]
        connection_ys = connection_object.GetAttribute('rY')
        valid_ys = [y for y in connection_ys if y != -1]

        xs.extend(valid_xs)
        ys.extend(valid_ys)

    logger.debug(f'xs: {get_formatted_fixed_space_numbers(xs)}')
    logger.debug(f'ys: {get_formatted_fixed_space_numbers(ys)}')

    # return a unique set of xs that match the above/below level, 0 exclusive
    return xs, ys


def get_relevant_xs(cubical, diagram, y_t, up_side):
    """Get the xs that are on the right side of the terminal"""

    use_up_side = up_side

    xs, ys = get_xs_ys(cubical, diagram)

    if use_up_side:
        above_xs = [x for x, y in zip(xs, ys) if y > y_t]
        logger.debug(f'Found {above_xs} that were above {y_t}')
        return above_xs
    else:
        below_xs = [x for x, y in zip(xs, ys) if y < y_t]
        logger.debug(f'Found {below_xs} that were below {y_t}')
        return below_xs


def get_formatted_fixed_space_numbers(nums, num_dec=5, width=10):
    form_nums = (f'{n:0.{num_dec}f}' for n in nums)
    return [f'{fn:>{width}}' for fn in form_nums]


def determine_term_x_obj_y(xt, yt, x1, y1, x2, y2):
    """
    return a tuple of x at the terminal spot and y at the object spot

    :param xt:  x term centre
    :param yt:  y term centre
    :param x1:
    :param y1:
    :param x2:
    :param y2:
    :return:
    """

    if yt == y1:
        x = x1
        y = y2
    elif yt == y2:
        x = x2
        y = y1
    else:
        dist_1 = dist_between_two_points(xt, yt, x1, y1)
        dist_2 = dist_between_two_points(xt, yt, x1, y1)
        if dist_1 < dist_2:
            x = x1
            y = y2
        else:
            x = x2
            y = y1

    return x, y


def dist_between_two_points(xa, ya, xb, yb):
    return ((xa - xb) ** 2 + (ya - yb) ** 2) ** 0.5


def get_centre_of_largest_gap(list_of_values):
    """
    for a numerical list of values return the midpoint of the largest gap
    between 2 sorted as consecutive numbers
    """
    logger.debug(list_of_values)
    list_of_values = list(sorted(set(list_of_values)))
    if len(list_of_values) == 1:
        list_of_values = list_of_values * 2
    logger.debug(list_of_values)

    differences = [x2 - x1 for (x1, x2) in zip(list_of_values, list_of_values[1:])]
    max_diff = max(differences)
    max_diff_index = differences.index(max_diff)
    ideal_x = (
        list_of_values[max_diff_index]
        + (list_of_values[max_diff_index + 1] - list_of_values[max_diff_index]) / 2
    )
    logger.debug(ideal_x)
    return ideal_x


if __name__ == '__main__':
    run_main()
