import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class AddController {

    @RequestMapping(value = "/add", method = RequestMethod.POST)
    public ModelAndView add(ModelAndView modelAndView) {
        modelAndView.addObject("message", "Item added successfully!");
        modelAndView.setViewName("add");
        return modelAndView;
    }

}